# Office Metadata Editor

<p align="center">
  <a href="README.md"><img src="https://img.shields.io/badge/Language-简体中文-grey?style=for-the-badge&logo=china" alt="简体中文"></a>
  <a href="README_EN.md"><img src="https://img.shields.io/badge/Language-English-blue?style=for-the-badge&logo=github" alt="English"></a>
</p>

<p align="center">
  <img src="assets/app_icon.png" width="128" height="128" alt="Office Metadata Editor Icon">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Release-v1.2.0-blue?style=flat-square&logo=github" alt="Release">
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Style-Win11%20Fluent%20Design-0078D4?style=flat-square&logo=windows11" alt="Style">
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/GUI-PySide6%20%2B%20FluentWidgets-purple?style=flat-square" alt="GUI">
</p>

A desktop editor for Microsoft Word, Excel and PowerPoint file metadata and Windows timestamps, built with **PySide6** and **Windows 11 Fluent Design**.

---

## v1.2.0: Excel and PowerPoint

| File type | Supported formats |
| --- | --- |
| Word documents | `.docx`, `.docm` |
| Excel workbooks | `.xlsx`, `.xlsm` |
| PowerPoint presentations | `.pptx`, `.pptm` |

- Browse or drop files to edit author, last modifier, title, description, company, creation/modification timestamps, total editing time and revision.
- Process mixed formats in one batch, scan folders recursively, apply fixed/random timestamps or anonymize metadata; skip Office `~$` lock files.
- Copy worksheets, formulas, charts, slides, images and VBA binary parts byte for byte. Create and register missing metadata parts when needed.
- Save through a temporary file in the same directory and atomic replacement. Reject corrupt or mismatched files; retain per-file batch results after completion.
- Fix batch creation timestamps and 64-bit Windows timestamp synchronization. Display local times and store UTC timestamps.

This tool edits file-level properties. It does not edit cells or slide content. Legacy `.doc` / `.xls` / `.ppt`, `.xlsb` and encrypted files are unsupported. Office applications differ in how they display and retain properties such as total editing time; Excel may not display it. Microsoft's [Excel built-in property documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.builtindocumentproperties) notes that some properties do not have defined values. Anonymization clears the listed metadata fields, not document content, annotations, tracked changes or custom properties.

v1.2.0 provides a portable Windows x64 EXE and ZIP, linked below, with a SHA-256 checksum attachment. Previous releases remain available.

---

## 📋 Release v1.1.0 Changelog

- 📝 **Comments / Description Editing**: Full support for viewing and editing Word comments metadata (`dc:description`) in single and batch modes, with auto-scrubbing during anonymization.
- 🎲 **Randomized Total Editing Time & Creation Date**: Specify customized min/max ranges for generating random editing durations and creation timestamps in batch operations.
- 🎨 **New 3D Win11 Fluent Design Icon**: Built-in 3D glassmorphism application logo, natively embedded into executables for window and taskbar rendering.
- 🛡️ **Privacy & Path Isolation**: Enhanced launch scripts and `.agents` workspace rule isolation to safeguard local private paths.
- ⚡ **Build Optimization**: Optimized PyInstaller build flags with asset bundling and exclusion of unused dependencies.

---

## 📥 Download Release (No Python Required)

No Python installation required. Download the portable standalone `.exe` for 64-bit Windows:

👉 **[View the v1.2.0 release](https://github.com/treetree-sha/WordMetadataEditor/releases/tag/v1.2.0)**

- [Download standalone EXE](https://github.com/treetree-sha/WordMetadataEditor/releases/download/v1.2.0/OfficeMetadataEditor_v1.2.0_x64.exe)
- [Download ZIP with instructions and license](https://github.com/treetree-sha/WordMetadataEditor/releases/download/v1.2.0/OfficeMetadataEditor_v1.2.0_x64.zip)
- [Download SHA-256 checksums](https://github.com/treetree-sha/WordMetadataEditor/releases/download/v1.2.0/OfficeMetadataEditor_v1.2.0_x64.SHA256.txt)

---

## 🌟 Key Features

1. **📄 Single File Fine Editing**
   * **Drag & Drop**: Drop any supported Word, Excel or PowerPoint file to inspect its metadata.
   * **Metadata Fields**:
     * **Creator (Author)** & **Last Modified By**
     * **Created Time** & **Modified Time** (includes a "Set to Now" shortcut)
     * **Total Editing Time** (in minutes, e.g. `120` = 2 hours)
     * **Revision Number**
     * Company Name & Document Title
   * **1-Click Anonymize**: Quickly clear author, modifier, company info, and reset editing time to zero.

2. **📂 Batch Processing & Privacy Cleaning**
   * Batch select mixed Word, Excel and PowerPoint files or scan entire directory trees recursively.
   * Asynchronous multi-threading ensures a responsive GUI with live progress bar and status table updates.
   * Apply unified author, modifier, editing time, or anonymize in bulk.

3. **⏰ Windows NTFS File System Timestamp Sync**
   * When enabled, modifying document XML metadata automatically syncs Windows NTFS file system **Created (ctime)** and **Modified (mtime)** timestamps via Win32 APIs.

4. **🎨 Win11 Fluent Design Interface**
   * **Dark Mode** and **Light Mode** theme switching.
   * **Dynamic Multi-Language Support** (Simplified Chinese & English).

5. **👤 Author & Open Source Links**
   * Integrated GitHub author profile and repository links in the Settings page.

---

## 📂 Project Structure

```
WordMetadataEditor/
├── metadata_engine.py      # Core engine (XML parsing & Win32 timestamp sync)
├── main.py                 # Main entry point (FluentWindow framework)
├── run.bat                 # Windows quick launch script
├── tests/                  # Office metadata, preservation and UI regression tests
├── README.md               # Chinese Documentation
├── README_EN.md            # English Documentation
└── ui/                     # UI Views
    ├── __init__.py
    ├── i18n.py                   # Internationalization engine (Chinese & English)
    ├── file_filters.py           # Shared Office file picker filters
    ├── single_file_interface.py  # Single file editor page
    ├── batch_interface.py        # Batch processing page
    └── settings_interface.py     # Settings page (Theme, Language & GitHub links)
```

---

## 🚀 Getting Started & Building

### 1. Run Locally
- **Method 1**: Double-click `run.bat` in Windows File Explorer.
- **Method 2**: Run `python main.py` in your terminal.

### 2. Verify Source

Install `pytest`, `python-docx`, `openpyxl` and `python-pptx` for tests. Application runtime dependencies remain `PySide6` and `PySide6-Fluent-Widgets`.

```powershell
python -m pytest tests -q
```

Tests cover six formats, byte preservation, missing metadata parts, failure atomicity, Windows timestamps, Qt forms and mixed batch processing. Macro tests use synthetic binary payloads to verify preservation without executing VBA.

### 3. Build Standalone EXE
Pack into a single executable using PyInstaller:
```bash
pyinstaller --noconsole --onefile --icon="assets/app_icon.ico" --add-data "assets;assets" --name "OfficeMetadataEditor_v1.2.0_x64" --exclude-module PyQt5 --clean main.py
```

---

## 📄 License

This project is licensed under the **[MIT License](LICENSE)**.

Developed with ❤️ by **[treetree-sha](https://github.com/treetree-sha)**.
