"""Windows structured storage backend. No Office automation or third-party runtime.

Uses IPropertySetStorage/IPropertyStorage on a caller-owned temporary copy.
The binary document streams are never parsed for editing or regenerated.
"""
import ctypes
import os
import struct
import uuid
from contextlib import contextmanager
from ctypes import wintypes as wt
from datetime import datetime, timedelta, timezone

from metadata_errors import MetadataError, EncryptedFileError


LEGACY_EXTENSIONS = ('.doc', '.xls', '.ppt')
CFB_SIGNATURE = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
EPOCH = datetime(1601, 1, 1, tzinfo=timezone.utc)
SUMMARY = 'F29F85E0-4FF9-1068-AB91-08002B27B3D9'
DOC_SUMMARY = 'D5CDD502-2E9C-101B-9397-08002B2CF9AE'
IID_STORAGE = '0000000B-0000-0000-C000-000000000046'
IID_PROPERTY_SETS = '0000013A-0000-0000-C000-000000000046'
FIELDS = {
    'title': 2, 'subject': 3, 'author': 4, 'comments': 6,
    'last_modified_by': 8, 'revision': 9, 'total_editing_time': 10,
    'created_time': 12, 'modified_time': 13, 'application': 18,
}
STGM_READ = 0x20  # READ | SHARE_DENY_WRITE
STGM_WRITE = 0x12  # READWRITE | SHARE_EXCLUSIVE
STG_E_FILENOTFOUND = 0x80030002
WINCALL = getattr(ctypes, 'WINFUNCTYPE', ctypes.CFUNCTYPE)


class GUID(ctypes.Structure):
    _fields_ = [('bytes', ctypes.c_ubyte * 16)]

    @classmethod
    def parse(cls, value):
        return cls.from_buffer_copy(uuid.UUID(value).bytes_le)


class _SpecValue(ctypes.Union):
    _fields_ = [('propid', wt.ULONG), ('name', wt.LPWSTR)]


class PROPSPEC(ctypes.Structure):
    _anonymous_ = ('value',)
    _fields_ = [('kind', wt.ULONG), ('value', _SpecValue)]


class _VariantValue(ctypes.Union):
    _fields_ = [('number', ctypes.c_ulonglong), ('integer', wt.LONG),
                ('short', ctypes.c_short), ('pointer', ctypes.c_void_p),
                # PROPVARIANT's largest union member on x64 is a count + pointer.
                ('storage', ctypes.c_ubyte * (16 if ctypes.sizeof(ctypes.c_void_p) == 8 else 8))]


class PROPVARIANT(ctypes.Structure):
    _anonymous_ = ('value',)
    _fields_ = [('vt', wt.WORD), ('reserved1', wt.WORD), ('reserved2', wt.WORD),
                ('reserved3', wt.WORD), ('value', _VariantValue)]


class STATSTG(ctypes.Structure):
    _fields_ = [('name', wt.LPWSTR), ('kind', wt.DWORD), ('size', ctypes.c_ulonglong),
                ('mtime', wt.FILETIME), ('ctime', wt.FILETIME), ('atime', wt.FILETIME),
                ('mode', wt.DWORD), ('locks', wt.DWORD), ('clsid', GUID),
                ('state', wt.DWORD), ('reserved', wt.DWORD)]


def _check(result, operation):
    if result < 0:
        raise MetadataError(f'{operation} failed (HRESULT 0x{result & 0xffffffff:08X})')


class Interface:
    def __init__(self, pointer):
        self.pointer = ctypes.c_void_p(pointer.value if isinstance(pointer, ctypes.c_void_p) else pointer)

    def call(self, index, argtypes=(), *args):
        table = ctypes.cast(self.pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        method = WINCALL(wt.LONG, ctypes.c_void_p, *argtypes)(table[index])
        return method(self.pointer, *args)

    def query(self, iid):
        output = ctypes.c_void_p()
        guid = GUID.parse(iid)
        _check(self.call(0, (ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)),
                         ctypes.byref(guid), ctypes.byref(output)), 'QueryInterface')
        return Interface(output)

    def close(self):
        if self.pointer.value:
            self.call(2)
            self.pointer = ctypes.c_void_p()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


@contextmanager
def compound_file(path, write=False):
    if os.name != 'nt':
        raise MetadataError('Legacy Office metadata requires Windows')
    ole = ctypes.WinDLL('ole32')
    ole.CoInitializeEx.argtypes = [ctypes.c_void_p, wt.DWORD]
    ole.CoInitializeEx.restype = wt.LONG
    result = ole.CoInitializeEx(None, 0)
    # A GUI thread may already be initialized as STA. Its apartment remains usable.
    initialized = result in (0, 1)
    if not initialized and (result & 0xffffffff) != 0x80010106:
        _check(result, 'CoInitializeEx')
    root = None
    try:
        ole.StgOpenStorageEx.argtypes = [wt.LPCWSTR, wt.DWORD, wt.DWORD, wt.DWORD,
                                      ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(GUID),
                                      ctypes.POINTER(ctypes.c_void_p)]
        ole.StgOpenStorageEx.restype = wt.LONG
        pointer = ctypes.c_void_p()
        guid = GUID.parse(IID_STORAGE)
        _check(ole.StgOpenStorageEx(os.path.abspath(path), STGM_WRITE if write else STGM_READ,
                                  5, 0, None, None, ctypes.byref(guid), ctypes.byref(pointer)),
               'Open compound file')
        root = Interface(pointer)
        yield root
    finally:
        if root is not None:
            root.close()
        if initialized:
            ole.CoUninitialize()


def entries(storage):
    pointer = ctypes.c_void_p()
    _check(storage.call(11, (wt.DWORD, ctypes.c_void_p, wt.DWORD, ctypes.POINTER(ctypes.c_void_p)),
                        0, None, 0, ctypes.byref(pointer)), 'Enumerate storage')
    ole = ctypes.WinDLL('ole32')
    ole.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    with Interface(pointer) as iterator:
        while True:
            item = STATSTG()
            fetched = wt.ULONG()
            result = iterator.call(3, (wt.ULONG, ctypes.POINTER(STATSTG), ctypes.POINTER(wt.ULONG)),
                                   1, ctypes.byref(item), ctypes.byref(fetched))
            _check(result, 'Read storage entry')
            if not fetched.value:
                break
            try:
                yield item.name, item.kind, item.size
            finally:
                # LPCWSTR field access converts to str; release the original COM pointer.
                raw = ctypes.cast(ctypes.byref(item), ctypes.POINTER(ctypes.c_void_p)).contents.value
                ole.CoTaskMemFree(raw)


@contextmanager
def open_stream(storage, name):
    pointer = ctypes.c_void_p()
    _check(storage.call(4, (wt.LPCWSTR, ctypes.c_void_p, wt.DWORD, wt.DWORD,
                           ctypes.POINTER(ctypes.c_void_p)), name, None, 0x10, 0, ctypes.byref(pointer)),
           'Open document stream')
    with Interface(pointer) as stream:
        yield stream


def read_bytes(stream, count):
    buffer = ctypes.create_string_buffer(count)
    actual = wt.ULONG()
    _check(stream.call(3, (ctypes.c_void_p, wt.ULONG, ctypes.POINTER(wt.ULONG)),
                       buffer, count, ctypes.byref(actual)), 'Read document stream')
    return buffer.raw[:actual.value]


def validate_document(storage, extension):
    names = {name: (kind, size) for name, kind, size in entries(storage)}
    if any(name in names for name in ('EncryptedPackage', 'EncryptionInfo', 'Encryption', 'EncryptedSummary')):
        raise EncryptedFileError('Encrypted Office files are not supported')
    if extension == '.doc':
        if 'WordDocument' not in names:
            raise MetadataError('Not a Word 97-2003 compound document')
        with open_stream(storage, 'WordDocument') as stream:
            header = read_bytes(stream, 32)
        if len(header) != 32 or struct.unpack_from('<H', header)[0] != 0xA5EC:
            raise MetadataError('Invalid Word document header')
        version, flags = struct.unpack_from('<H', header, 2)[0], struct.unpack_from('<H', header, 10)[0]
        if flags & 0x0100:
            raise EncryptedFileError('Encrypted Word files are not supported')
        if version not in (0xC1, 0xD9, 0x101, 0x10C, 0x112):
            raise MetadataError('Unsupported Word binary format version')
        table = '1Table' if flags & 0x0200 else '0Table'
        if table not in names:
            raise MetadataError('Word table stream is missing')
    elif extension == '.xls':
        name = 'Workbook' if 'Workbook' in names else 'Book'
        if name not in names:
            raise MetadataError('Not an Excel 97-2003 compound workbook')
        with open_stream(storage, name) as stream:
            first = True
            while True:
                header = read_bytes(stream, 4)
                if len(header) != 4:
                    raise MetadataError('Incomplete Excel workbook globals')
                record, size = struct.unpack('<HH', header)
                data = read_bytes(stream, size)
                if len(data) != size:
                    raise MetadataError('Truncated Excel record')
                if first:
                    if record != 0x0809 or len(data) < 4 or struct.unpack_from('<HH', data) != (0x0600, 0x0005):
                        raise MetadataError('Only Excel BIFF8 (97-2003) workbooks are supported')
                    first = False
                if record == 0x002F:
                    raise EncryptedFileError('Encrypted Excel files are not supported')
                if record == 0x000A:
                    break
    elif extension == '.ppt':
        if not {'PowerPoint Document', 'Current User'}.issubset(names):
            raise MetadataError('Not a PowerPoint 97-2003 compound presentation')
        with open_stream(storage, 'Current User') as stream:
            header = read_bytes(stream, 20)
        if len(header) < 20 or struct.unpack_from('<H', header, 2)[0] != 0x0FF6:
            raise MetadataError('Invalid PowerPoint CurrentUserAtom')
        token = struct.unpack_from('<I', header, 12)[0]
        if token == 0xF3D1C4DF:
            raise EncryptedFileError('Encrypted PowerPoint files are not supported')
        if token != 0xE391C05F:
            raise MetadataError('Unsupported PowerPoint binary format')
    else:
        raise MetadataError('File extension does not match the compound document')


@contextmanager
def property_set(sets, fmtid, write=False):
    pointer = ctypes.c_void_p()
    guid = GUID.parse(fmtid)
    mode = STGM_WRITE if write else 0x10
    result = sets.call(4, (ctypes.POINTER(GUID), wt.DWORD, ctypes.POINTER(ctypes.c_void_p)),
                       ctypes.byref(guid), mode, ctypes.byref(pointer))
    if (result & 0xffffffff) == STG_E_FILENOTFOUND:
        if not write:
            yield None
            return
        _check(sets.call(3, (ctypes.POINTER(GUID), ctypes.c_void_p, wt.DWORD, wt.DWORD,
                            ctypes.POINTER(ctypes.c_void_p)), ctypes.byref(guid), None, 0, mode,
                         ctypes.byref(pointer)), 'Create property set')
    else:
        _check(result, 'Open property set')
    with Interface(pointer) as properties:
        yield properties


def read_property(properties, identifier):
    spec = PROPSPEC(kind=1, propid=identifier)
    value = PROPVARIANT()
    try:
        _check(properties.call(3, (wt.ULONG, ctypes.POINTER(PROPSPEC), ctypes.POINTER(PROPVARIANT)),
                               1, ctypes.byref(spec), ctypes.byref(value)), 'Read property')
        if value.vt in (0, 1):
            return None
        if value.vt in (2, 18):
            return value.short & 0xffff
        if value.vt in (3, 19):
            return value.integer
        if value.vt == 64:
            return value.number
        if value.vt in (8, 31):
            return ctypes.wstring_at(value.pointer) if value.pointer else ''
        if value.vt == 30:
            return ctypes.string_at(value.pointer) if value.pointer else b''
        raise MetadataError(f'Unexpected type for property {identifier}: {value.vt}')
    finally:
        ole = ctypes.WinDLL('ole32')
        ole.PropVariantClear.argtypes = [ctypes.POINTER(PROPVARIANT)]
        ole.PropVariantClear(ctypes.byref(value))


def _text(value, codepage):
    if isinstance(value, bytes):
        encoding = 'mbcs' if codepage in (None, 1200) else f'cp{codepage}'
        try:
            return value.decode(encoding)
        except (LookupError, UnicodeError) as exc:
            raise MetadataError(f'Cannot decode OLE property code page {codepage}') from exc
    return '' if value is None else str(value)


def read_metadata(path, extension=None):
    extension = extension or os.path.splitext(os.fspath(path))[1].lower()
    result = {}
    with compound_file(path) as storage:
        validate_document(storage, extension)
        with storage.query(IID_PROPERTY_SETS) as sets:
            with property_set(sets, SUMMARY) as props:
                if props is not None:
                    cp = read_property(props, 1)
                    for key, identifier in FIELDS.items():
                        value = read_property(props, identifier)
                        if value is None:
                            continue
                        if key == 'total_editing_time':
                            result[key] = str(value // 600000000)
                        elif key in ('created_time', 'modified_time'):
                            result[key] = (EPOCH + timedelta(microseconds=value // 10)).isoformat().replace('+00:00', 'Z') if value else ''
                        else:
                            result[key] = _text(value, cp)
            with property_set(sets, DOC_SUMMARY) as props:
                if props is not None:
                    result['company'] = _text(read_property(props, 15), read_property(props, 1))
    return result


def _write_property(properties, identifier, value, numeric=False):
    spec = PROPSPEC(kind=1, propid=identifier)
    variant = PROPVARIANT()
    if numeric:
        if not 0 <= value < 2**64:
            raise MetadataError('OLE FILETIME value is out of range')
        variant.vt = 64  # FILETIME, including edit durations in 100 ns units.
        variant.number = value
    else:
        text = str(value)
        if '\0' in text:
            raise MetadataError('Property text cannot contain null characters')
        buffer = ctypes.create_unicode_buffer(text)
        variant.vt = 31  # Explicit Unicode, independent of the existing ANSI code page.
        variant.pointer = ctypes.cast(buffer, ctypes.c_void_p)
    _check(properties.call(4, (wt.ULONG, ctypes.POINTER(PROPSPEC), ctypes.POINTER(PROPVARIANT), wt.ULONG),
                           1, ctypes.byref(spec), ctypes.byref(variant), 2), 'Write property')


def write_metadata_to_copy(path, extension, metadata):
    """Caller owns this copy and atomically replaces the original after success."""
    with compound_file(path, write=True) as storage:
        validate_document(storage, extension)
        with storage.query(IID_PROPERTY_SETS) as sets:
            selected = {key: value for key, value in metadata.items()
                        if key in FIELDS and key != 'application' and value is not None}
            if selected:
                with property_set(sets, SUMMARY, write=True) as props:
                    for key, value in selected.items():
                        numeric = key in ('total_editing_time', 'created_time', 'modified_time')
                        if key == 'total_editing_time':
                            value = int(value) * 600000000
                        elif key in ('created_time', 'modified_time'):
                            if not value:
                                continue
                            dt = datetime.fromisoformat(value.replace('Z', '+00:00'))
                            delta = dt.astimezone(timezone.utc) - EPOCH
                            value = (delta.days * 86400 + delta.seconds) * 10000000 + delta.microseconds * 10
                            if not 0 <= value < 2**64:
                                raise MetadataError('OLE timestamps must be on or after 1601-01-01')
                        _write_property(props, FIELDS[key], value, numeric)
                    _check(props.call(9, (wt.DWORD,), 0), 'Commit properties')
            if metadata.get('company') is not None:
                with property_set(sets, DOC_SUMMARY, write=True) as props:
                    _write_property(props, 15, metadata['company'])
                    _check(props.call(9, (wt.DWORD,), 0), 'Commit company')
        _check(storage.call(9, (wt.DWORD,), 0), 'Commit compound file')
