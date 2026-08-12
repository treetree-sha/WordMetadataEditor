import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import ctypes
from ctypes import wintypes

# XML Namespaces
NAMESPACES = {
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
}

# Register namespaces to prevent ns0: prefixes when writing XML
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


def set_windows_file_times(filepath: str, created_time: datetime = None, modified_time: datetime = None, access_time: datetime = None):
    """Update Windows file system Creation, Last Write, and Last Access times using Win32 API."""
    if not os.path.exists(filepath):
        return False
    
    try:
        def datetime_to_filetime(dt: datetime):
            if dt is None:
                return None
            timestamp = dt.timestamp()
            filetime_val = int((timestamp + 11644473600) * 10000000)
            return wintypes.FILETIME(filetime_val & 0xFFFFFFFF, filetime_val >> 32)

        ft_create = datetime_to_filetime(created_time)
        ft_access = datetime_to_filetime(access_time or modified_time)
        ft_write = datetime_to_filetime(modified_time)

        GENERIC_WRITE = 0x40000000
        FILE_SHARE_READ = 0x00000001
        FILE_SHARE_WRITE = 0x00000002
        OPEN_EXISTING = 3
        FILE_FLAG_BACKUP_SEMANTICS = 0x02000000

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.CreateFileW(
            filepath,
            GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None,
            OPEN_EXISTING,
            FILE_FLAG_BACKUP_SEMANTICS,
            None
        )

        if handle == -1 or handle == 0:
            return False

        p_create = ctypes.byref(ft_create) if ft_create else None
        p_access = ctypes.byref(ft_access) if ft_access else None
        p_write = ctypes.byref(ft_write) if ft_write else None

        result = kernel32.SetFileTime(handle, p_create, p_access, p_write)
        kernel32.CloseHandle(handle)
        return bool(result)
    except Exception as e:
        print(f"Error setting file times: {e}")
        return False


class WordMetadataEngine:
    """Engine to inspect and update Word (.docx) document metadata."""

    @staticmethod
    def read_metadata(file_path: str) -> dict:
        """Read all relevant metadata from a .docx file."""
        metadata = {
            'author': '',
            'last_modified_by': '',
            'created_time': '',
            'modified_time': '',
            'revision': '1',
            'title': '',
            'subject': '',
            'total_editing_time': '0', # In minutes
            'company': '',
            'application': '',
            # File system info
            'fs_created_time': '',
            'fs_modified_time': '',
            'file_size': 0
        }

        if not os.path.exists(file_path) or not file_path.lower().endswith('.docx'):
            return metadata

        # File system info
        stat = os.stat(file_path)
        metadata['file_size'] = stat.st_size
        metadata['fs_created_time'] = datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
        metadata['fs_modified_time'] = datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                # Read docProps/core.xml
                if 'docProps/core.xml' in zf.namelist():
                    core_xml = zf.read('docProps/core.xml')
                    root = ET.fromstring(core_xml)
                    
                    creator_node = root.find('dc:creator', NAMESPACES)
                    if creator_node is not None and creator_node.text:
                        metadata['author'] = creator_node.text
                    
                    last_mod_node = root.find('cp:lastModifiedBy', NAMESPACES)
                    if last_mod_node is not None and last_mod_node.text:
                        metadata['last_modified_by'] = last_mod_node.text

                    created_node = root.find('dcterms:created', NAMESPACES)
                    if created_node is not None and created_node.text:
                        metadata['created_time'] = created_node.text

                    modified_node = root.find('dcterms:modified', NAMESPACES)
                    if modified_node is not None and modified_node.text:
                        metadata['modified_time'] = modified_node.text

                    rev_node = root.find('cp:revision', NAMESPACES)
                    if rev_node is not None and rev_node.text:
                        metadata['revision'] = rev_node.text

                    title_node = root.find('dc:title', NAMESPACES)
                    if title_node is not None and title_node.text:
                        metadata['title'] = title_node.text

                    subject_node = root.find('dc:subject', NAMESPACES)
                    if subject_node is not None and subject_node.text:
                        metadata['subject'] = subject_node.text

                # Read docProps/app.xml
                if 'docProps/app.xml' in zf.namelist():
                    app_xml = zf.read('docProps/app.xml')
                    root = ET.fromstring(app_xml)

                    total_time_node = root.find('ep:TotalTime', NAMESPACES)
                    if total_time_node is None:
                        for elem in root.iter():
                            if elem.tag.endswith('TotalTime'):
                                total_time_node = elem
                                break

                    if total_time_node is not None and total_time_node.text:
                        metadata['total_editing_time'] = total_time_node.text

                    company_node = root.find('ep:Company', NAMESPACES)
                    if company_node is None:
                        for elem in root.iter():
                            if elem.tag.endswith('Company'):
                                company_node = elem
                                break
                    if company_node is not None and company_node.text:
                        metadata['company'] = company_node.text

                    app_node = root.find('ep:Application', NAMESPACES)
                    if app_node is None:
                        for elem in root.iter():
                            if elem.tag.endswith('Application'):
                                app_node = elem
                                break
                    if app_node is not None and app_node.text:
                        metadata['application'] = app_node.text

        except Exception as e:
            print(f"Error reading docx metadata: {e}")

        return metadata

    @staticmethod
    def write_metadata(file_path: str, new_meta: dict, sync_fs_time: bool = True) -> bool:
        """Write updated metadata back into the .docx ZIP container."""
        if not os.path.exists(file_path) or not file_path.lower().endswith('.docx'):
            return False

        temp_dir = tempfile.mkdtemp()
        temp_docx = os.path.join(temp_dir, 'modified.docx')

        try:
            with zipfile.ZipFile(file_path, 'r') as zin, zipfile.ZipFile(temp_docx, 'w') as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)

                    if item.filename == 'docProps/core.xml':
                        content = WordMetadataEngine._update_core_xml(content, new_meta)
                    elif item.filename == 'docProps/app.xml':
                        content = WordMetadataEngine._update_app_xml(content, new_meta)

                    zout.writestr(item, content)

            # Replace original file safely
            shutil.move(temp_docx, file_path)
            shutil.rmtree(temp_dir, ignore_errors=True)

            # Sync file system time if requested
            if sync_fs_time:
                created_dt = None
                modified_dt = None

                if new_meta.get('created_time'):
                    try:
                        clean_time = new_meta['created_time'].replace('Z', '+00:00')
                        created_dt = datetime.fromisoformat(clean_time)
                    except Exception:
                        pass
                if new_meta.get('modified_time'):
                    try:
                        clean_time = new_meta['modified_time'].replace('Z', '+00:00')
                        modified_dt = datetime.fromisoformat(clean_time)
                    except Exception:
                        pass

                if created_dt or modified_dt:
                    set_windows_file_times(file_path, created_time=created_dt, modified_time=modified_dt)

            return True

        except Exception as e:
            print(f"Error writing docx metadata: {e}")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
            return False

    @staticmethod
    def _update_core_xml(xml_bytes: bytes, new_meta: dict) -> bytes:
        """Helper to modify core.xml tree."""
        root = ET.fromstring(xml_bytes)

        def set_or_create(tag_with_ns, value, attrib=None):
            prefix, tag_name = tag_with_ns.split(':')
            ns_uri = NAMESPACES[prefix]
            elem = root.find(f'{prefix}:{tag_name}', NAMESPACES)
            if elem is None:
                elem = ET.SubElement(root, f'{{{ns_uri}}}{tag_name}')
                if attrib:
                    elem.attrib.update(attrib)
            if value is not None:
                elem.text = str(value)

        if 'author' in new_meta:
            set_or_create('dc:creator', new_meta['author'])
        if 'last_modified_by' in new_meta:
            set_or_create('cp:lastModifiedBy', new_meta['last_modified_by'])
        if 'created_time' in new_meta and new_meta['created_time']:
            set_or_create('dcterms:created', new_meta['created_time'], {'{http://www.w3.org/2001/XMLSchema-instance}type': 'dcterms:W3CDTF'})
        if 'modified_time' in new_meta and new_meta['modified_time']:
            set_or_create('dcterms:modified', new_meta['modified_time'], {'{http://www.w3.org/2001/XMLSchema-instance}type': 'dcterms:W3CDTF'})
        if 'revision' in new_meta:
            set_or_create('cp:revision', new_meta['revision'])
        if 'title' in new_meta:
            set_or_create('dc:title', new_meta['title'])
        if 'subject' in new_meta:
            set_or_create('dc:subject', new_meta['subject'])

        return ET.tostring(root, encoding='utf-8', xml_declaration=True)

    @staticmethod
    def _update_app_xml(xml_bytes: bytes, new_meta: dict) -> bytes:
        """Helper to modify app.xml tree."""
        root = ET.fromstring(xml_bytes)
        ep_ns = NAMESPACES['ep']

        if 'total_editing_time' in new_meta:
            elem = root.find('ep:TotalTime', NAMESPACES)
            if elem is None:
                for child in root:
                    if child.tag.endswith('TotalTime'):
                        elem = child
                        break
            if elem is None:
                elem = ET.SubElement(root, f'{{{ep_ns}}}TotalTime')
            elem.text = str(new_meta['total_editing_time'])

        if 'company' in new_meta:
            elem = root.find('ep:Company', NAMESPACES)
            if elem is None:
                for child in root:
                    if child.tag.endswith('Company'):
                        elem = child
                        break
            if elem is None:
                elem = ET.SubElement(root, f'{{{ep_ns}}}Company')
            elem.text = str(new_meta['company'])

        return ET.tostring(root, encoding='utf-8', xml_declaration=True)
