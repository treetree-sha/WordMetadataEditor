"""Optional developer check using independently opened Microsoft Office files.

Production metadata editing never imports win32com or starts Office. This tool
requires Office and pywin32, creates its own application instances, and only
opens generated test copies. All generated files remain under test_documents.
"""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from metadata_engine import OfficeMetadataEngine


def application(progid):
    import pythoncom
    from win32com.client import dynamic
    return dynamic.Dispatch(pythoncom.CoCreateInstance(
        progid, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch))


def cleanup(app, document, powerpoint=False):
    try:
        if document is not None:
            if powerpoint:
                document.Close()
            else:
                document.Close(False)
    finally:
        if app is not None:
            app.Quit()


def generate(folder):
    if any((folder / ('sample' + ext)).exists() for ext in ('.doc', '.xls', '.xlsb', '.ppt')):
        raise RuntimeError('Sample files already exist. Choose an empty --fixtures directory.')
    app = document = None
    try:
        app = application('Word.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        document = app.Documents.Add()
        document.Content.Text = 'Legacy Word fixture.\r中文正文保持不变。\r'
        document.SaveAs2(str(folder / 'sample.doc'), FileFormat=0)
    finally:
        cleanup(app, document)
    app = document = None
    try:
        app = application('Excel.Application')
        app.Visible = False
        app.DisplayAlerts = False
        document = app.Workbooks.Add()
        sheet = document.Worksheets(1)
        sheet.Name = '统计'
        sheet.Range('A1').Value = '项目'
        sheet.Range('B1').Value = 3
        sheet.Range('A2').Value = '中文数据'
        sheet.Range('B2').Value = 7
        sheet.Range('B3').Formula = '=SUM(B1:B2)'
        sheet.Range('B3').Font.Bold = True
        for extension, file_format in (('.xls', 56), ('.xlsb', 50)):
            document.SaveAs(str(folder / ('sample' + extension)), FileFormat=file_format)
        sheet = None
    finally:
        cleanup(app, document)
    app = document = None
    try:
        app = application('PowerPoint.Application')
        document = app.Presentations.Add(False)
        slide = document.Slides.Add(1, 12)
        textbox = slide.Shapes.AddTextbox(1, 50, 50, 500, 100)
        textbox.TextFrame.TextRange.Text = 'Legacy PowerPoint fixture / 中文幻灯片'
        document.SaveAs(str(folder / 'sample.ppt'), 1)
        textbox = slide = None
    finally:
        cleanup(app, document, powerpoint=True)


def open_and_verify(path, expected):
    extension = path.suffix
    app = document = properties = None
    progid = 'Word.Application' if extension == '.doc' else (
        'PowerPoint.Application' if extension == '.ppt' else 'Excel.Application')
    try:
        app = application(progid)
        app.AutomationSecurity = 3  # Force-disable macros during optional application checks.
        if extension == '.doc':
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Open(str(path), ConfirmConversions=False,
                                          ReadOnly=True, AddToRecentFiles=False)
            assert 'Legacy Word fixture.' in document.Content.Text
        elif extension in ('.xls', '.xlsb'):
            app.Visible = False
            app.DisplayAlerts = False
            document = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
            assert document.Worksheets(1).Range('B3').Formula == '=SUM(B1:B2)'
            assert document.Worksheets(1).Range('B3').Value == 10
            assert document.Worksheets(1).Range('B3').Font.Bold
        else:
            document = app.Presentations.Open(str(path), ReadOnly=True, WithWindow=False)
            assert document.Slides.Count == 1
            assert 'Legacy PowerPoint fixture' in document.Slides(1).Shapes(1).TextFrame.TextRange.Text
        properties = document.BuiltInDocumentProperties
        for name, key in (('Author', 'author'), ('Title', 'title'),
                          ('Comments', 'comments'), ('Company', 'company')):
            assert properties(name).Value == expected[key], f'{extension}: {name}'
        displayed_time = str(properties('Total Editing Time').Value)
        return {'content_verified': True, 'unicode_properties_verified': True,
                'summary_editing_minutes': expected['total_editing_time'],
                'application_editing_minutes': displayed_time}
    finally:
        properties = None
        cleanup(app, document, powerpoint=(extension == '.ppt'))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--fixtures', type=Path, default=ROOT / 'test_documents/native_office')
    parser.add_argument('--generate', action='store_true', help='Create new samples in an empty directory')
    args = parser.parse_args()
    folder = args.fixtures.resolve()
    folder.mkdir(parents=True, exist_ok=True)
    if args.generate:
        generate(folder)
    output = Path(tempfile.mkdtemp(prefix='office-check-', dir=folder))
    updates = dict(author='中文作者测试', last_modified_by='修改者', title='中文测试标题',
                   comments='备注 <test> & 中文', company='测试公司', revision='7',
                   total_editing_time='123', created_time='2025-01-02T03:04:05Z',
                   modified_time='2026-03-04T05:06:07Z')
    result = {}
    for extension in ('.doc', '.xls', '.xlsb', '.ppt'):
        source = folder / ('sample' + extension)
        original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
        target = output / source.name
        shutil.copyfile(source, target)
        assert OfficeMetadataEngine.write_metadata(target, updates, sync_fs_time=False)
        readback = OfficeMetadataEngine.read_metadata(target)
        assert all(readback[key] == value for key, value in updates.items())
        result[extension] = open_and_verify(target, updates)
        assert hashlib.sha256(source.read_bytes()).hexdigest() == original_hash
        print(extension + ' passed', flush=True)
    (output / 'native_check.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(str(output / 'native_check.json'), flush=True)


if __name__ == '__main__':
    main()
