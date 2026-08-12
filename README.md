# Word 文档属性高级编辑器 (Word Metadata Editor)

<p align="center">
  <a href="README.md"><img src="https://img.shields.io/badge/Language-简体中文-blue?style=for-the-badge&logo=china" alt="简体中文"></a>
  <a href="README_EN.md"><img src="https://img.shields.io/badge/Language-English-grey?style=for-the-badge&logo=github" alt="English"></a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Release-v1.0.0-blue?style=flat-square&logo=github" alt="Release">
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Style-Win11%20Fluent%20Design-0078D4?style=flat-square&logo=windows11" alt="Style">
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/GUI-PySide6%20%2B%20FluentWidgets-purple?style=flat-square" alt="GUI">
</p>

一套基于 **PySide6** 与 **Windows 11 Fluent Design** 打造的高颜值 Word (`.docx`) 文档元数据及文件系统属性高级修改软件。

---

## 📥 绿色版快速下载 (Release Download)

无需安装 Python 或配置任何依赖环境，直接下载 Windows 独立单文件版：

👉 **[点击前往 GitHub Releases 下载绿色单文件版 (v1.0.0)](https://github.com/treetree-sha/WordMetadataEditor/releases)**

---

## 🌟 核心功能特性

1. **📄 单文件精细编辑 (Single File Editor)**
   * **拖拽支持**：拖拽任意 `.docx` 文件至窗口即可自动解析并加载所有元数据。
   * **核心属性修改**：
     * **作者 (Creator)** & **最后修改者 (Last Modified By)**
     * **创建时间 (Created Time)** & **修改时间 (Modified Time)**（提供“一键设为当前时间”）
     * **总编辑时间 (Total Editing Time)**（单位：分钟，如设为 `120` 表示 2 小时）
     * **修订版本号 (Revision)**
     * 公司名称 (Company) & 文档标题 (Title)
   * **一键脱敏**：快速清空作者、修改人、公司名称，并将总编辑时间归零。

2. **📂 批量处理与脱敏 (Batch Processing)**
   * 支持同时选择多个 Word 文档或扫描整个文件夹（递归子目录）。
   * 采用后台多线程处理（UI 界面顺滑不卡顿），支持实时进度条显示与表格状态更新。
   * 支持统一批量设置作者、修改人、编辑时间或一键清理隐私。

3. **⏰ Windows 文件系统时间同步联动 (NTFS Sync)**
   * 勾选“同步更新系统文件时间”后，软件在修改 `.docx` 内部 XML 的同时，通过 Win32 API 自动修改 Windows NTFS 文件系统的 **【创建时间 (ctime)】** 与 **【修改时间 (mtime)】**。

4. **🎨 Windows 11 Fluent Design 高颜值界面**
   * 支持 **暗黑模式 (Dark Mode)** 和 **浅色模式 (Light Mode)** 无缝切换。
   * 支持 **简体中文 / English 界面语言无缝切换**。

5. **👤 作者信息与 GitHub 联动**
   * 设置页面内置作者 GitHub 个人主页与项目开源仓库链接。

---

## 📂 项目结构

```
WordMetadataEditor/
├── metadata_engine.py      # 元数据核心引擎 (解析/修改 core.xml、app.xml 及 Win32 API 时间同步)
├── main.py                 # 主程序入口 (FluentWindow 导航框架)
├── run.bat                 # Windows 一键启动脚本
├── test_document.docx      # 默认测试文档
├── README.md               # 中文说明文档
├── README_EN.md            # English Documentation
└── ui/                     # UI 视图模块
    ├── __init__.py
    ├── i18n.py                   # 国际化多语言引擎 (Simplified Chinese & English)
    ├── single_file_interface.py  # 单文件精修页面
    ├── batch_interface.py        # 批量处理与脱敏页面
    └── settings_interface.py     # 设置与关于页面 (包含语言切换与 GitHub 作者链接)
```

---

## 🚀 启动与打包指南

### 1. 本地启动运行
- **方法一**：在 Windows 资源管理器中双击 **`run.bat`**。
- **方法二**：命令行运行 `python main.py` 或 `D:\Programs\Anaconda\python.exe main.py`。

### 2. 打包为 Windows 单文件 EXE
使用 PyInstaller 进行打包：
```bash
pyinstaller --noconsole --onefile --name "WordMetadataEditor" --exclude-module PyQt5 --clean main.py
```

---

## 📄 开源许可证

本项目基于 **[MIT License](LICENSE)** 开源。

Developed with ❤️ by **[treetree-sha](https://github.com/treetree-sha)**.
