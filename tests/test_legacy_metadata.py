"""OLE property tests use independent pywin32 fixtures, never Office automation.

Real Office-generated samples may additionally be placed in
test_documents/native_office (or OFFICE_NATIVE_FIXTURES); originals are copied.
"""
import hashlib
import os
import shutil
import struct
import uuid
from pathlib import Path
from unittest.mock import patch

import pytest

from metadata_engine import OfficeMetadataEngine as Engine, MetadataError, EncryptedFileError

pytestmark = pytest.mark.skipif(os.name != 'nt', reason='Windows structured storage')


def _com():
    return pytest.importorskip('pythoncom')


def ole_contents(path):
    """Independent IStorage reader compares every stream, including nested VBA."""
    com = _com()
    root = com.StgOpenStorage(str(path), None, 0x20)
    result = {}
    def visit(storage, prefix=''):
        for entry in storage.EnumElements():
            name, kind = entry[:2]
            key = prefix + name
            if kind == 1:
                child = storage.OpenStorage(name, None, 0x10)
                visit(child, key + '/')
                child = None
            elif kind == 2:
                stream = storage.OpenStream(name, None, 0x10, 0)
                chunks = []
                while True:
                    chunk = stream.Read(65536)
                    if not chunk:
                        break
                    chunks.append(chunk)
                result[key] = b''.join(chunks)
                stream = None
    visit(root)
    root = None
    return result


def _ansi_summary():
    # Real OLEPS layout with CP1252, a non-ASCII ANSI author, and an unknown field.
    author = 'André'.encode('cp1252') + b'\0'
    values = [(1, struct.pack('<Ih2x', 2, 1252)),
              (4, struct.pack('<II', 30, len(author)) + author),
              (100, struct.pack('<Ii', 3, 777))]
    payload = b''
    offsets = []
    base = 8 + len(values) * 8
    for identifier, value in values:
        offsets.append(struct.pack('<II', identifier, base + len(payload)))
        payload += value + b'\0' * (-len(value) % 4)
    section = struct.pack('<II', base + len(payload), len(values)) + b''.join(offsets) + payload
    return (struct.pack('<HHI16sI', 0xFFFE, 0, 0x00020006, bytes(16), 1)
            + uuid.UUID('F29F85E0-4FF9-1068-AB91-08002B27B3D9').bytes_le
            + struct.pack('<I', 48) + section)


@pytest.fixture
def make_legacy(tmp_path):
    com = _com()
    def make(ext, encrypted=False, properties=True):
        path = tmp_path / ('旧版属性测试' + ext)
        storage = com.StgCreateDocfile(str(path), 0x1012, 0)
        def stream(name, data, parent=storage):
            handle = parent.CreateStream(name, 0x1012, 0, 0)
            handle.Write(data)
            handle = None
        if ext == '.doc':
            header = bytearray(32)
            struct.pack_into('<HH', header, 0, 0xA5EC, 0xC1)
            struct.pack_into('<H', header, 10, 0x0100 if encrypted else 0)
            stream('WordDocument', bytes(header) + bytes(range(256)) * 40)
            stream('0Table', b'opaque word table data')
        elif ext == '.xls':
            payload = struct.pack('<HHHH', 0x0809, 4, 0x0600, 0x0005)
            if encrypted:
                payload += struct.pack('<HHH', 0x002F, 2, 0)
            payload += struct.pack('<HH', 0x000A, 0)
            stream('Workbook', payload + b'opaque worksheet/formula data' * 300)
        elif ext == '.ppt':
            header = struct.pack('<HHIIIHH', 0, 0x0FF6, 20, 20,
                                 0xF3D1C4DF if encrypted else 0xE391C05F, 0, 0)
            stream('Current User', header)
            stream('PowerPoint Document', b'opaque slide data' * 700)
        nested = storage.CreateStorage('VBA', 0x1012, 0, 0)
        stream('Module1', bytes(range(256)) * 20, nested)
        nested.Commit(0)
        nested = None
        stream('UnknownStream', b'custom non-metadata bytes')
        if properties:
            stream('\x05SummaryInformation', _ansi_summary())
            sets = storage.QueryInterface(com.IID_IPropertySetStorage)
            props = sets.Create(com.FMTID_DocSummaryInformation, com.IID_NULL, 0, 0x12)
            props.WriteMultiple((14, 15), ('Keep Manager', 'Old Company'))
            props.Commit(0)
            props = None
            custom_id = com.MakeIID('{D5CDD505-2E9C-101B-9397-08002B2CF9AE}')
            custom = sets.Create(custom_id, com.IID_NULL, 0, 0x12)
            custom.WriteMultiple(('KeepCustom',), ('自定义属性保持',))
            custom.Commit(0)
            custom = None
            sets = None
        storage.Commit(0)
        storage = None
        return path
    return make


@pytest.mark.parametrize('ext', ['.doc', '.xls', '.ppt'])
def test_unicode_metadata_and_every_document_stream_preserved(make_legacy, ext):
    path = make_legacy(ext)
    assert Engine.read_metadata(path)['author'] == 'André'
    before = ole_contents(path)
    updates = dict(author='中文作者', last_modified_by='最后修改者', company='测试单位',
                   title='标题 & <测试>', subject='主题', comments='多行备注\n第二行' * 100,
                   revision='8', total_editing_time='120',
                   created_time='2025-01-02T03:04:05Z', modified_time='2026-02-03T04:05:06Z')
    assert Engine.write_metadata(path, updates, False)
    actual = Engine.read_metadata(path)
    assert all(actual[key] == value for key, value in updates.items())
    after = ole_contents(path)
    excluded = {'\x05SummaryInformation', '\x05DocumentSummaryInformation'}
    assert {key: val for key, val in before.items() if key not in excluded} == {
        key: val for key, val in after.items() if key not in excluded}
    com = _com()
    root = com.StgOpenStorage(str(path), None, 0x20)
    sets = root.QueryInterface(com.IID_IPropertySetStorage)
    props = sets.Open(com.FMTID_SummaryInformation, 0x10)
    assert props.ReadMultiple((100,)) == (777,)
    props = None
    props = sets.Open(com.FMTID_DocSummaryInformation, 0x10)
    assert props.ReadMultiple((14,)) == ('Keep Manager',)
    props = None
    props = sets.Open(com.MakeIID('{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'), 0x10)
    assert props.ReadMultiple(('KeepCustom',)) == ('自定义属性保持',)
    props = sets = root = None


@pytest.mark.parametrize('ext', ['.doc', '.xls', '.ppt'])
def test_missing_property_sets_are_created(make_legacy, ext):
    path = make_legacy(ext, properties=False)
    before = ole_contents(path)
    assert Engine.read_metadata(path)['author'] == ''
    assert Engine.write_metadata(path, {'author': '新增作者', 'company': '新增单位', 'total_editing_time': '15'}, False)
    assert Engine.read_metadata(path)['author'] == '新增作者'
    after = ole_contents(path)
    assert all(after[name] == data for name, data in before.items())


@pytest.mark.parametrize('ext', ['.doc', '.xls', '.ppt'])
def test_encrypted_legacy_files_are_rejected_without_changes(make_legacy, ext):
    path = make_legacy(ext, encrypted=True)
    original = path.read_bytes()
    with pytest.raises(EncryptedFileError):
        Engine.read_metadata(path)
    assert not Engine.write_metadata(path, {'author': 'Do not change'}, False)
    assert path.read_bytes() == original


def test_encrypted_xlsb_package_is_rejected(tmp_path):
    com = _com()
    path = tmp_path / 'encrypted.xlsb'
    root = com.StgCreateDocfile(str(path), 0x1012, 0)
    for name in ('EncryptionInfo', 'EncryptedPackage'):
        stream = root.CreateStream(name, 0x1012, 0, 0)
        stream.Write(b'encrypted package placeholder')
        stream = None
    root.Commit(0)
    root = None
    original = path.read_bytes()
    with pytest.raises(EncryptedFileError):
        Engine.read_metadata(path)
    assert not Engine.write_metadata(path, {'author': 'Do not change'}, False)
    assert path.read_bytes() == original


def test_legacy_write_failure_is_atomic(make_legacy):
    path = make_legacy('.doc')
    original = path.read_bytes()
    with patch('metadata_engine.os.replace', side_effect=PermissionError('locked')):
        assert not Engine.write_metadata(path, {'author': 'Attempt'}, False)
    assert path.read_bytes() == original
    assert not list(path.parent.glob('.office_metadata_*.tmp'))
    assert not Engine.write_metadata(path, {'author': 'bad\0text'}, False)
    assert path.read_bytes() == original


def test_legacy_extension_mismatch_is_rejected(make_legacy):
    path = make_legacy('.doc')
    path = path.rename(path.with_suffix('.xls'))
    original = path.read_bytes()
    with pytest.raises(MetadataError):
        Engine.read_metadata(path)
    assert not Engine.write_metadata(path, {'author': 'Attempt'}, False)
    assert path.read_bytes() == original


@pytest.mark.parametrize('ext', ['.doc', '.xls', '.ppt', '.xlsb'])
def test_native_office_fixture_round_trip(tmp_path, ext):
    folder = Path(os.environ.get('OFFICE_NATIVE_FIXTURES',
                  str(Path(__file__).resolve().parents[1] / 'test_documents/native_office')))
    source = folder / ('sample' + ext)
    if not source.is_file():
        pytest.skip('Optional real Office-generated fixture is unavailable')
    original_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    path = tmp_path / source.name
    shutil.copyfile(source, path)
    if ext == '.xlsb':
        from test_office_metadata import contents
        before = contents(path)
    else:
        before = ole_contents(path)
    assert Engine.write_metadata(path, {'author': '本机真实样本测试', 'company': '兼容性测试', 'total_editing_time': '75'}, False)
    assert Engine.read_metadata(path)['author'] == '本机真实样本测试'
    after = contents(path) if ext == '.xlsb' else ole_contents(path)
    excluded = {'\x05SummaryInformation', '\x05DocumentSummaryInformation', 'docProps/core.xml', 'docProps/app.xml'}
    assert {key: val for key, val in before.items() if key not in excluded} == {
        key: val for key, val in after.items() if key not in excluded}
    assert hashlib.sha256(source.read_bytes()).hexdigest() == original_hash
