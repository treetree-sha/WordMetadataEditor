# Word Metadata Editor

<p align="center">
  <a href="README.md"><img src="https://img.shields.io/badge/Language-简体中文-grey?style=for-the-badge&logo=china" alt="简体中文"></a>
  <a href="README_EN.md"><img src="https://img.shields.io/badge/Language-English-blue?style=for-the-badge&logo=github" alt="English"></a>
</p>

<p align="center">
  <img src="assets/app_icon.png" width="128" height="128" alt="Word Metadata Editor Icon">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Release-v1.1.0-blue?style=flat-square&logo=github" alt="Release">
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Style-Win11%20Fluent%20Design-0078D4?style=flat-square&logo=windows11" alt="Style">
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/GUI-PySide6%20%2B%20FluentWidgets-purple?style=flat-square" alt="GUI">
</p>

A modern, high-aesthetic desktop tool for modifying Word (`.docx`) metadata & Windows file system attributes, built with **PySide6** and **Windows 11 Fluent Design**.

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

👉 **[Download Standalone Package (v1.1.0)](https://github.com/treetree-sha/WordMetadataEditor/releases)**

---

## 🌟 Key Features

1. **📄 Single File Fine Editing**
   * **Drag & Drop**: Simply drop any `.docx` file into the window to automatically inspect all metadata.
   * **Metadata Fields**:
     * **Creator (Author)** & **Last Modified By**
     * **Created Time** & **Modified Time** (includes a "Set to Now" shortcut)
     * **Total Editing Time** (in minutes, e.g. `120` = 2 hours)
     * **Revision Number**
     * Company Name & Document Title
   * **1-Click Anonymize**: Quickly clear author, modifier, company info, and reset editing time to zero.

2. **📂 Batch Processing & Privacy Cleaning**
   * Batch select multiple Word files or scan entire directory trees recursively.
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
├── test_document.docx      # Sample test file
├── README.md               # Chinese Documentation
├── README_EN.md            # English Documentation
└── ui/                     # UI Views
    ├── __init__.py
    ├── i18n.py                   # Internationalization engine (Chinese & English)
    ├── single_file_interface.py  # Single file editor page
    ├── batch_interface.py        # Batch processing page
    └── settings_interface.py     # Settings page (Theme, Language & GitHub links)
```

---

## 🚀 Getting Started & Building

### 1. Run Locally
- **Method 1**: Double-click `run.bat` in Windows File Explorer.
- **Method 2**: Run `python main.py` in your terminal.

### 2. Build Standalone EXE
Pack into a single executable using PyInstaller:
```bash
pyinstaller --noconsole --onefile --icon="assets/app_icon.ico" --add-data "assets;assets" --name "WordMetadataEditor" --exclude-module PyQt5 --clean main.py
```

---

## 📄 License

This project is licensed under the **[MIT License](LICENSE)**.

Developed with ❤️ by **[treetree-sha](https://github.com/treetree-sha)**.
