"""Real OOXML round trips, content preservation, and failure atomicity."""
import os
import shutil
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path
from unittest.mock import patch

import pytest
from docx import Document
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Inches

from metadata_engine import (
    OfficeMetadataEngine as Engine, WordMetadataEngine, MetadataError,
    NAMESPACES, OFFICE_FORMATS, PROPERTY_PARTS, format_ooxml_datetime,
    is_supported_file,
)


def contents(path):
    with zipfile.ZipFile(path) as zf:
        assert zf.testzip() is None
        return {item.filename: zf.read(item) for item in zf.infolist()}


def rewrite(path, parts):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)


@pytest.fixture
def make_office(tmp_path):
    def make(ext):
        path = tmp_path / ('属性测试' + ext)
        family, main, content_type = OFFICE_FORMATS[ext]
        if ext == '.xlsb':
            native_dir = Path(os.environ.get('OFFICE_NATIVE_FIXTURES',
                              str(Path(__file__).resolve().parents[1] / 'test_documents/native_office')))
            sample = native_dir / 'sample.xlsb'
            if sample.is_file():
                shutil.copyfile(sample, path)
            else:
                # Container-level fixture only; native Office checks are opt-in.
                parts = {
                    main: b'opaque binary workbook test payload',
                    '[Content_Types].xml': f'<Types xmlns="{NAMESPACES["ct"]}"><Default Extension="bin" ContentType="{content_type}"/></Types>'.encode(),
                    '_rels/.rels': f'<Relationships xmlns="{NAMESPACES["rel"]}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{main}"/></Relationships>'.encode(),
                    'docProps/core.xml': f'<cp:coreProperties xmlns:cp="{NAMESPACES["cp"]}"/>'.encode(),
                    'docProps/app.xml': f'<Properties xmlns="{NAMESPACES["ep"]}"/>'.encode(),
                }
                # Register existing metadata just as a normal Office package does.
                types = ET.fromstring(parts['[Content_Types].xml'])
                rels = ET.fromstring(parts['_rels/.rels'])
                for index, (name, _, _, rel_type, ctype) in enumerate(PROPERTY_PARTS.values(), 2):
                    ET.SubElement(types, f'{{{NAMESPACES["ct"]}}}Override', PartName='/' + name, ContentType=ctype)
                    ET.SubElement(rels, f'{{{NAMESPACES["rel"]}}}Relationship', Id=f'rId{index}', Type=rel_type, Target=name)
                parts['[Content_Types].xml'] = ET.tostring(types)
                parts['_rels/.rels'] = ET.tostring(rels)
                rewrite(path, parts)
            return path
        if family == 'Word':
            doc = Document()
            doc.add_heading('Original content', 0)
            doc.add_paragraph('正文与格式应保留')
            doc.save(path)
        elif family == 'Excel':
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = '统计'
            sheet.append(['项目', '金额'])
            sheet.append(['A', 12])
            sheet.append(['B', 30])
            sheet['B4'] = '=SUM(B2:B3)'
            sheet['B4'].font = Font(bold=True, color='FF0000')
            sheet.merge_cells('D1:F1')
            workbook.create_sheet('附表')
            chart = BarChart()
            chart.add_data(Reference(sheet, min_col=2, min_row=1, max_row=3), titles_from_data=True)
            sheet.add_chart(chart, 'H1')
            workbook.save(path)
        else:
            presentation = Presentation()
            slide = presentation.slides.add_slide(presentation.slide_layouts[0])
            slide.shapes.title.text = '原始演示内容'
            slide.shapes.add_picture(str(Path(__file__).resolve().parents[1] / 'assets/app_icon.png'), Inches(1), Inches(2), width=Inches(1))
            slide.notes_slide.notes_text_frame.text = '备注页内容'
            presentation.save(path)
        if ext.endswith('m'):
            # Synthetic macro payload tests byte preservation without executing VBA.
            parts = contents(path)
            types = ET.fromstring(parts['[Content_Types].xml'])
            for node in types:
                if node.get('PartName') == '/' + main:
                    node.set('ContentType', content_type)
            ET.SubElement(types, f'{{{NAMESPACES["ct"]}}}Default',
                          Extension='bin', ContentType='application/vnd.ms-office.vbaProject')
            parts['[Content_Types].xml'] = ET.tostring(types)
            directory, filename = main.rsplit('/', 1)
            rel_path = directory + '/_rels/' + filename + '.rels'
            rels = ET.fromstring(parts[rel_path])
            ET.SubElement(rels, f'{{{NAMESPACES["rel"]}}}Relationship', Id='rIdTestMacro',
                          Type='http://schemas.microsoft.com/office/2006/relationships/vbaProject', Target='vbaProject.bin')
            parts[rel_path] = ET.tostring(rels)
            parts[directory + '/vbaProject.bin'] = bytes(range(256)) * 3
            rewrite(path, parts)
        return path
    return make


@pytest.mark.parametrize('ext', OFFICE_FORMATS)
def test_round_trip_preserves_every_non_metadata_part(make_office, ext):
    path = make_office(ext)
    before = contents(path)
    updates = dict(author='新作者 & 团队', last_modified_by='修改者', title='新标题',
                   subject='新主题', comments='说明 <保留>\n第二行', company='研发部',
                   total_editing_time='123', revision='9',
                   created_time='2025-01-02T11:04:05+08:00', modified_time='2026-03-04T05:06:07Z')
    assert Engine.write_metadata(path, updates, sync_fs_time=False)
    meta = Engine.read_metadata(path)
    for key, value in updates.items():
        assert meta[key] == (format_ooxml_datetime(value) if key.endswith('_time') and key != 'total_editing_time' else value)
    after = contents(path)
    assert before.keys() == after.keys()
    assert {key: data for key, data in before.items() if key not in ('docProps/core.xml', 'docProps/app.xml')} == {
        key: data for key, data in after.items() if key not in ('docProps/core.xml', 'docProps/app.xml')}
    if ext == '.xlsx':
        wb = load_workbook(path)
        assert wb['统计']['B4'].value == '=SUM(B2:B3)'
        assert wb['统计']['B4'].font.bold
        assert wb.sheetnames == ['统计', '附表']
        assert len(wb['统计']._charts) == 1
        wb.close()
    elif ext == '.pptx':
        presentation = Presentation(path)
        assert presentation.slides[0].shapes.title.text == '原始演示内容'
        assert presentation.slides[0].notes_slide.notes_text_frame.text == '备注页内容'
    elif ext == '.docx':
        assert Document(path).paragraphs[1].text == '正文与格式应保留'


@pytest.mark.parametrize('ext', ['.docx', '.xlsx', '.pptx'])
def test_missing_properties_are_created_and_registered(make_office, ext):
    path = make_office(ext)
    parts = contents(path)
    for kind, (name, _, _, rel_type, _) in PROPERTY_PARTS.items():
        parts.pop(name, None)
    types = ET.fromstring(parts['[Content_Types].xml'])
    rels = ET.fromstring(parts['_rels/.rels'])
    for node in list(types):
        if node.get('PartName', '').startswith('/docProps/'):
            types.remove(node)
    for node in list(rels):
        if node.get('Type') in [info[3] for info in PROPERTY_PARTS.values()]:
            rels.remove(node)
    parts['[Content_Types].xml'] = ET.tostring(types)
    parts['_rels/.rels'] = ET.tostring(rels)
    rewrite(path, parts)
    assert Engine.write_metadata(path, {'author': 'Alice', 'company': 'Acme', 'total_editing_time': '45'}, False)
    assert Engine.read_metadata(path)['author'] == 'Alice'
    assert Engine.read_metadata(path)['company'] == 'Acme'
    after = contents(path)
    ids = [node.get('Id') for node in ET.fromstring(after['_rels/.rels'])]
    assert len(ids) == len(set(ids))
    for kind, (name, _, _, rel_type, content_type) in PROPERTY_PARTS.items():
        assert any(node.get('PartName') == '/' + name and node.get('ContentType') == content_type
                   for node in ET.fromstring(after['[Content_Types].xml']))
        assert any(node.get('Type') == rel_type and node.get('Target') == name
                   for node in ET.fromstring(after['_rels/.rels']))
    for name in parts.keys() - {'[Content_Types].xml', '_rels/.rels'}:
        assert after[name] == parts[name]


def test_selected_patch_does_not_change_other_metadata(make_office):
    path = make_office('.xlsx')
    before = contents(path)
    assert Engine.write_metadata(path, {'author': 'Only change author'}, False)
    after = contents(path)
    assert after['docProps/app.xml'] == before['docProps/app.xml']
    assert Engine.read_metadata(path)['comments'] == ''


@pytest.mark.parametrize('invalid', ['invalid_date', '2025-99-01', '2025-02-30T12:00:00Z'])
def test_invalid_timestamp_leaves_original_intact(make_office, invalid):
    path = make_office('.pptx')
    original = path.read_bytes()
    assert not Engine.write_metadata(path, {'created_time': invalid}, False)
    assert path.read_bytes() == original


def test_failed_replace_preserves_original_and_cleans_temp(make_office):
    path = make_office('.xlsx')
    original = path.read_bytes()
    with patch('metadata_engine.os.replace', side_effect=PermissionError('locked')):
        assert not Engine.write_metadata(path, {'author': 'Alice'}, False)
    assert path.read_bytes() == original
    assert not list(path.parent.glob('.office_metadata_*.tmp'))


@pytest.mark.parametrize('kind', ['renamed', 'corrupt', 'legacy', 'bad_xml', 'bad_app_xml', 'wrong_property_root'])
def test_invalid_files_do_not_report_success(make_office, kind):
    path = make_office('.xlsx')
    if kind == 'renamed':
        path = path.rename(path.with_suffix('.pptx'))
    elif kind == 'legacy':
        path = path.rename(path.with_suffix('.xls'))
    elif kind == 'corrupt':
        path.write_bytes(b'not a zip file')
    else:
        parts = contents(path)
        if kind == 'bad_app_xml':
            parts['docProps/app.xml'] = b'<broken'
        elif kind == 'wrong_property_root':
            parts['docProps/core.xml'] = b'<unrelated/>'
        else:
            parts['docProps/core.xml'] = b'<broken'
        rewrite(path, parts)
    original = path.read_bytes()
    with pytest.raises(MetadataError):
        Engine.read_metadata(path)
    assert not Engine.write_metadata(path, {'author': 'Alice'}, False)
    assert path.read_bytes() == original


def test_utc_conversion_and_compatibility():
    assert WordMetadataEngine is Engine
    assert format_ooxml_datetime('2025-01-02T11:04:05+08:00') == '2025-01-02T03:04:05Z'
    assert format_ooxml_datetime('2025-01-02T03:04:05.123456Z') == '2025-01-02T03:04:05.123456Z'
    assert is_supported_file('报告.XLSX')
    assert not is_supported_file('~$报告.xlsx')
    assert is_supported_file('旧版.ppt')


@pytest.mark.skipif(os.name != 'nt', reason='Windows file system API')
def test_real_windows_timestamps(make_office):
    path = make_office('.xlsx')
    created = datetime(2025, 1, 2, 3, 4, 5, tzinfo=timezone.utc)
    modified = created + timedelta(days=3)
    assert Engine.write_metadata(path, {'created_time': created.isoformat(), 'modified_time': modified.isoformat()}, True)
    stat = path.stat()
    assert abs(stat.st_ctime - created.timestamp()) < 1
    assert abs(stat.st_mtime - modified.timestamp()) < 1
