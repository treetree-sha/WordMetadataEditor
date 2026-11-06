import ctypes
import os
import posixpath
import shutil
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from ctypes import wintypes
from datetime import datetime, timezone

from metadata_errors import MetadataError, EncryptedFileError
from ole_metadata import (LEGACY_EXTENSIONS, CFB_SIGNATURE, compound_file, entries,
                          read_metadata as read_legacy_metadata, write_metadata_to_copy)


NAMESPACES = {
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

# OPC metadata is shared; document content is copied verbatim.
OFFICE_FORMATS = {
    '.docx': ('Word', 'word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
    '.docm': ('Word', 'word/document.xml', 'application/vnd.ms-word.document.macroEnabled.main+xml'),
    '.xlsx': ('Excel', 'xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
    '.xlsm': ('Excel', 'xl/workbook.xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'),
    '.xlsb': ('Excel', 'xl/workbook.bin', 'application/vnd.ms-excel.sheet.binary.macroEnabled.main'),
    '.pptx': ('PowerPoint', 'ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'),
    '.pptm': ('PowerPoint', 'ppt/presentation.xml', 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'),
}
SUPPORTED_EXTENSIONS = tuple(OFFICE_FORMATS) + LEGACY_EXTENSIONS
FILE_PATTERNS = ' '.join('*' + ext for ext in SUPPORTED_EXTENSIONS)
CORE_FIELDS = {
    'author': 'dc:creator', 'last_modified_by': 'cp:lastModifiedBy',
    'created_time': 'dcterms:created', 'modified_time': 'dcterms:modified',
    'revision': 'cp:revision', 'title': 'dc:title', 'subject': 'dc:subject',
    'comments': 'dc:description',
}
APP_FIELDS = {'total_editing_time': 'ep:TotalTime', 'company': 'ep:Company', 'application': 'ep:Application'}
PROPERTY_PARTS = {
    'core': ('docProps/core.xml', NAMESPACES['cp'], 'coreProperties',
             'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
             'application/vnd.openxmlformats-package.core-properties+xml'),
    'app': ('docProps/app.xml', NAMESPACES['ep'], 'Properties',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
            'application/vnd.openxmlformats-officedocument.extended-properties+xml'),
}


def is_supported_file(file_path):
    path = os.fspath(file_path)
    return (os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS
            and not os.path.basename(path).startswith('~$'))


def format_ooxml_datetime(value):
    """Convert local naive or offset-aware input to UTC."""
    dt = value if isinstance(value, datetime) else datetime.fromisoformat(value.replace('Z', '+00:00'))
    return dt.astimezone(timezone.utc).isoformat().replace('+00:00', 'Z')


def set_windows_file_times(filepath, created_time=None, modified_time=None, access_time=None):
    """Set NTFS times with pointer-sized Win32 handles, including on 64-bit."""
    if os.name != 'nt' or not os.path.isfile(filepath):
        return False
    try:
        def to_filetime(dt):
            if dt is None:
                return None
            value = int((dt.timestamp() + 11644473600) * 10000000)
            return wintypes.FILETIME(value & 0xFFFFFFFF, value >> 32)

        kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                                        wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        kernel32.CreateFileW.restype = wintypes.HANDLE
        kernel32.SetFileTime.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 3
        kernel32.SetFileTime.restype = wintypes.BOOL
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL
        handle = kernel32.CreateFileW(os.path.abspath(filepath), 0x100, 0x3, None, 3, 0x80, None)
        if handle in (None, ctypes.c_void_p(-1).value):
            return False
        try:
            times = [to_filetime(dt) for dt in (created_time, access_time, modified_time)]
            return bool(kernel32.SetFileTime(handle, *[ctypes.byref(ft) if ft else None for ft in times]))
        finally:
            kernel32.CloseHandle(handle)
    except (OSError, ValueError, OverflowError):
        return False


class OfficeMetadataEngine:
    """Read/write Word, Excel and PowerPoint metadata without resaving content."""

    @staticmethod
    def _inspect_package(zf, file_path):
        if not is_supported_file(file_path) or os.path.splitext(file_path)[1].lower() not in OFFICE_FORMATS:
            raise MetadataError('Unsupported Office file format')
        names = zf.namelist()
        if len(names) != len(set(names)):
            raise MetadataError('Duplicate package entries')
        _, main_part, content_type = OFFICE_FORMATS[os.path.splitext(file_path)[1].lower()]
        types = ET.fromstring(zf.read('[Content_Types].xml'))
        rels = ET.fromstring(zf.read('_rels/.rels'))
        overrides = {node.get('PartName'): node.get('ContentType') for node in types}
        defaults = {node.get('Extension', '').lower(): node.get('ContentType') for node in types
                    if node.tag == f'{{{NAMESPACES["ct"]}}}Default'}
        main_type = overrides.get('/' + main_part, defaults.get(main_part.rsplit('.', 1)[-1].lower()))
        if main_part not in names or main_type != content_type:
            raise MetadataError('File extension does not match the Office package')
        paths = {}
        for kind, (default, ns, root_tag, rel_type, _) in PROPERTY_PARTS.items():
            related = [node for node in rels if node.get('Type') == rel_type]
            if len(related) > 1:
                raise MetadataError('Duplicate property relationships')
            target = default
            if related:
                if related[0].get('TargetMode') == 'External':
                    raise MetadataError('External property relationships are unsupported')
                target = posixpath.normpath(related[0].get('Target', '').lstrip('/'))
                if target in ('', '.', '..') or target.startswith('../'):
                    raise MetadataError('Invalid property part path')
            if target in names:
                root = ET.fromstring(zf.read(target))
                if root.tag != f'{{{ns}}}{root_tag}':
                    raise MetadataError('Invalid property XML root')
            paths[kind] = target
        if len(set(paths.values())) != len(paths) or any(
                path in (main_part, '[Content_Types].xml', '_rels/.rels') for path in paths.values()):
            raise MetadataError('Invalid property part path')
        return types, rels, paths

    @staticmethod
    def _reject_encrypted_package(file_path):
        with open(file_path, 'rb') as source:
            signature = source.read(8)
        if signature == CFB_SIGNATURE:
            with compound_file(file_path) as storage:
                names = {name for name, _, _ in entries(storage)}
                if 'EncryptedPackage' in names or 'EncryptionInfo' in names:
                    raise EncryptedFileError('Encrypted Office packages are not supported')
            raise MetadataError('File extension does not match the Office package')

    @staticmethod
    def read_metadata(file_path):
        file_path = os.fspath(file_path)
        try:
            if not is_supported_file(file_path):
                raise MetadataError('Unsupported Office file format')
            metadata = {key: '' for key in (*CORE_FIELDS, *APP_FIELDS)}
            metadata.update(revision='1', total_editing_time='0')
            if os.path.splitext(file_path)[1].lower() in LEGACY_EXTENSIONS:
                metadata.update(read_legacy_metadata(file_path))
            else:
                OfficeMetadataEngine._reject_encrypted_package(file_path)
                with zipfile.ZipFile(file_path, 'r') as zf:
                    _, _, paths = OfficeMetadataEngine._inspect_package(zf, file_path)
                    for kind, fields in (('core', CORE_FIELDS), ('app', APP_FIELDS)):
                        if paths[kind] in zf.namelist():
                            root = ET.fromstring(zf.read(paths[kind]))
                            for key, tag in fields.items():
                                node = root.find(tag, NAMESPACES)
                                if node is not None and node.text:
                                    metadata[key] = node.text
            stat = os.stat(file_path)
            metadata.update(
                file_size=stat.st_size,
                fs_created_time=datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                fs_modified_time=datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
            )
            return metadata
        except MetadataError:
            raise
        except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile, RuntimeError, ValueError, OverflowError) as exc:
            raise MetadataError(f'Cannot read Office metadata: {exc}') from exc

    @staticmethod
    def write_metadata(file_path, new_meta, sync_fs_time=True):
        """Update a same-directory copy and replace only after all writes succeed."""
        file_path = os.path.abspath(os.fspath(file_path))
        temp_path = None
        try:
            if not is_supported_file(file_path):
                raise MetadataError('Unsupported Office file format')
            extension = os.path.splitext(file_path)[1].lower()
            new_meta = dict(new_meta)
            for key in ('created_time', 'modified_time'):
                if new_meta.get(key):
                    new_meta[key] = format_ooxml_datetime(new_meta[key])
            for key in ('total_editing_time', 'revision'):
                if key in new_meta and int(new_meta[key]) < 0:
                    raise ValueError(f'{key} must be nonnegative')
            if extension in LEGACY_EXTENSIONS:
                read_legacy_metadata(file_path)  # Reject encrypted/mismatched sources before copying.
            else:
                OfficeMetadataEngine._reject_encrypted_package(file_path)
            fd, temp_path = tempfile.mkstemp(prefix='.office_metadata_', suffix='.tmp', dir=os.path.dirname(file_path))
            os.close(fd)
            if extension in LEGACY_EXTENSIONS:
                shutil.copyfile(file_path, temp_path)
                write_metadata_to_copy(temp_path, extension, new_meta)
                actual = read_legacy_metadata(temp_path, extension=extension)
                editable = set(CORE_FIELDS) | {'company', 'total_editing_time'}
                for key, value in new_meta.items():
                    if key in editable and value is not None:
                        if key in ('created_time', 'modified_time') and not value:
                            continue
                        if actual.get(key, '') != str(value):
                            raise MetadataError(f'Legacy property did not round-trip: {key}')
            elif not OfficeMetadataEngine._write_package(file_path, temp_path, new_meta):
                return True
            # Windows scanners/indexers can briefly hold a newly created file.
            for attempt in range(4):
                try:
                    os.replace(temp_path, file_path)
                    break
                except OSError as exc:
                    if getattr(exc, 'winerror', None) not in (5, 32, 33) or attempt == 3:
                        raise
                    time.sleep(0.05 * (attempt + 1))
            if sync_fs_time:
                times = {key: datetime.fromisoformat(new_meta[key].replace('Z', '+00:00'))
                         for key in ('created_time', 'modified_time') if new_meta.get(key)}
                if times and not set_windows_file_times(file_path, times.get('created_time'), times.get('modified_time')):
                    print('Metadata saved, but file system timestamps could not be synchronized.')
            return True
        except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile, RuntimeError, ValueError, TypeError, OverflowError) as exc:
            print(f'Error writing Office metadata: {exc}')
            return False
        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)

    @staticmethod
    def _write_package(file_path, temp_path, new_meta):
        with zipfile.ZipFile(file_path, 'r') as zin:
            types, rels, paths = OfficeMetadataEngine._inspect_package(zin, file_path)
            updates = {}
            manifests_changed = False
            for kind, fields, updater in (
                ('core', CORE_FIELDS, OfficeMetadataEngine._update_core_xml),
                ('app', {'total_editing_time', 'company'}, OfficeMetadataEngine._update_app_xml),
            ):
                if not any(key in new_meta for key in fields):
                    continue
                path = paths[kind]
                _, ns, root_tag, rel_type, content_type = PROPERTY_PARTS[kind]
                xml = zin.read(path) if path in zin.namelist() else ET.tostring(ET.Element(f'{{{ns}}}{root_tag}'))
                updates[path] = updater(xml, new_meta)
                if not any(node.get('PartName') == '/' + path for node in types):
                    ET.SubElement(types, f'{{{NAMESPACES["ct"]}}}Override',
                                  PartName='/' + path, ContentType=content_type)
                    manifests_changed = True
                if not any(node.get('Type') == rel_type for node in rels):
                    used_ids = {node.get('Id') for node in rels}
                    index = 1
                    while f'rId{index}' in used_ids:
                        index += 1
                    ET.SubElement(rels, f'{{{NAMESPACES["rel"]}}}Relationship',
                                  Id=f'rId{index}', Type=rel_type, Target=path)
                    manifests_changed = True
            if manifests_changed:
                updates['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
                updates['_rels/.rels'] = ET.tostring(rels, encoding='utf-8', xml_declaration=True)
            if not updates:
                return False
            with zipfile.ZipFile(temp_path, 'w') as zout:
                zout.comment = zin.comment
                for item in zin.infolist():
                    content = updates.pop(item.filename) if item.filename in updates else zin.read(item)
                    zout.writestr(item, content)
                for path, content in updates.items():
                    zout.writestr(path, content, compress_type=zipfile.ZIP_DEFLATED)
        return True

    @staticmethod
    def _update_core_xml(xml_bytes, new_meta):
        root = ET.fromstring(xml_bytes)
        for key, tag in CORE_FIELDS.items():
            if key not in new_meta or new_meta[key] is None:
                continue
            if key in ('created_time', 'modified_time') and not new_meta[key]:
                continue
            elem = root.find(tag, NAMESPACES)
            if elem is None:
                prefix, name = tag.split(':')
                elem = ET.SubElement(root, f'{{{NAMESPACES[prefix]}}}{name}')
            if key in ('created_time', 'modified_time'):
                elem.set(f'{{{NAMESPACES["xsi"]}}}type', 'dcterms:W3CDTF')
            elem.text = str(new_meta[key])
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)

    @staticmethod
    def _update_app_xml(xml_bytes, new_meta):
        root = ET.fromstring(xml_bytes)
        for key in ('total_editing_time', 'company'):
            if key in new_meta and new_meta[key] is not None:
                elem = root.find(APP_FIELDS[key], NAMESPACES)
                if elem is None:
                    elem = ET.SubElement(root, f'{{{NAMESPACES["ep"]}}}{APP_FIELDS[key].split(":")[1]}')
                elem.text = str(new_meta[key])
        return ET.tostring(root, encoding='utf-8', xml_declaration=True)


# Keep existing scripts and integrations working.
WordMetadataEngine = OfficeMetadataEngine
