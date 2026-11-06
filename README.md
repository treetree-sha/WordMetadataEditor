# Office 文档属性高级编辑器 (Office Metadata Editor)

<p align="center">
  <a href="README.md"><img src="https://img.shields.io/badge/Language-简体中文-blue?style=for-the-badge&logo=china" alt="简体中文"></a>
  <a href="README_EN.md"><img src="https://img.shields.io/badge/Language-English-grey?style=for-the-badge&logo=github" alt="English"></a>
</p>

<p align="center">
  <img src="assets/app_icon.png" width="128" height="128" alt="Office Metadata Editor Icon">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Release-v1.3.0-blue?style=flat-square&logo=github" alt="Release">
  <img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/Style-Win11%20Fluent%20Design-0078D4?style=flat-square&logo=windows11" alt="Style">
  <img src="https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/GUI-PySide6%20%2B%20FluentWidgets-purple?style=flat-square" alt="GUI">
</p>

一套基于 **PySide6** 与 **Windows 11 Fluent Design** 打造的 Office 文件元数据及文件系统属性编辑软件，支持 Microsoft Word、Excel 和 PowerPoint。

---

## v1.3.0：XLSB 与旧版 Office 文件

| 文件类型 | 支持格式 |
| --- | --- |
| Word 文档 | `.docx`、`.docm`、常见未加密 `.doc` |
| Excel 工作簿 | `.xlsx`、`.xlsm`、`.xlsb`、未加密 BIFF8 `.xls`（97–2003） |
| PowerPoint 演示文稿 | `.pptx`、`.pptm`、常见未加密 `.ppt`（97–2003） |

- 新格式已接入单文件浏览、拖拽、混合批量处理、递归扫描、属性脱敏及 Windows 文件时间同步。
- `.xlsb` 直接修改包内属性，保留二进制工作簿、公式和其他部件；兼容通过 Default 或 Override 声明主部件内容类型的文件。
- `.doc` / `.xls` / `.ppt` 在 Windows 上通过系统 OLE 属性接口读写；无需安装 Office，也无需增加第三方运行依赖。采用临时副本、属性回读验证及原子替换保存。
- 更新作者、修改者、标题、备注、公司、时间、修订号等摘要属性，同时保留未知属性、自定义属性及非属性流。加密文件、错误扩展名和不支持的旧格式会被拒绝修改。
- 不保证所有应用都显示或保留相同统计值。实测旧版 `.doc` 的 Word 界面可能采用内部统计而不采用摘要中的编辑时长；界面会显示此限制。正文、修订、批注和云端历史不属于属性脱敏范围。

v1.3.0 Windows x64 绿色版已包含上述格式支持，下载入口与 SHA-256 校验文件见下方；历史版本继续保留。

---

## v1.2.0：Excel 与 PowerPoint 支持

| 文件类型 | 支持格式 |
| --- | --- |
| Word 文档 | `.docx`、`.docm` |
| Excel 工作簿 | `.xlsx`、`.xlsm` |
| PowerPoint 演示文稿 | `.pptx`、`.pptm` |

- **单文件编辑**：浏览、拖拽、查看及修改作者、最后修改者、标题、备注、公司、创建/修改时间、总编辑时间与修订号。
- **混合批量处理**：三类文件可加入同一任务，支持递归扫描文件夹、统一属性、随机时间范围和属性脱敏，自动跳过 `~$` 临时文件。
- **内容保留**：仅更新文件属性；工作表、公式、图表、幻灯片、图片与 VBA 二进制内容按原始字节复制。缺失的属性文件会补建并注册到包中。
- **保存与结果反馈**：同目录临时文件及原子替换；损坏或格式不匹配的文件报告失败，批量任务结束后保留每个文件的处理结果。
- **时间修复**：修复批量创建时间未生效；界面使用本地时间，文件属性存为 UTC；修复 64 位 Windows 文件时间同步。

范围：本工具修改的是整个文件的属性，不编辑单元格或幻灯片正文。v1.2.0 发行版不支持旧版 `.doc` / `.xls` / `.ppt`、`.xlsb` 和加密文件。总编辑时间等属性在不同 Office 应用中的显示及保留方式可能不同；Excel 可能不显示该字段。Microsoft 的 [Excel 内置文档属性说明](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.builtindocumentproperties) 也指出，并非每个内置属性都有已定义的值。脱敏处理清理此工具列出的属性，不清理正文、批注、修订记录或自定义属性。

当前 v1.3.0 下载包含本节功能与上方新增的旧格式支持；此处保留 v1.2.0 功能说明作为版本历史。

---

## 📋 Release v1.1.0 更新日志 (Changelog)

- 📝 **新增【备注 / 摘要说明 (Comments)】编辑**：单文件精细编辑与批量处理模式全面支持查看、统一修改 Word 文档备注（`dc:description`）信息，脱敏时支持一键清空。
- 🎲 **新增【随机编辑总时长范围】与【随机创建时间范围】**：批量修改模式下新增在指定区间内（如 30~180 分钟）随机生成总编辑时间与创建时间。
- 🎨 **全新 Win11 3D Fluent Design 软件图标**：内置专属设计的 3D 立体悬浮图标，打包为 `.exe` 文件后窗口与任务栏完整原生呈现。
- 🛡️ **本地隐私安全隔离**：优化脚本启动配置与路径保护，`.agents` 配置与本地开发路径已安全隔离。
- ⚡ **打包体积与启动优化**：优化 PyInstaller 打包构建流程，剔除冗余依赖，软件体积更小且启动迅速。

---

## 📥 绿色版快速下载 (Release Download)

无需安装 Python 或配置任何依赖环境，直接下载 Windows 独立单文件版：

👉 **[查看 v1.3.0 发行版](https://github.com/treetree-sha/WordMetadataEditor/releases/tag/v1.3.0)**

- [下载 EXE 单文件版](https://github.com/treetree-sha/WordMetadataEditor/releases/download/v1.3.0/OfficeMetadataEditor_v1.3.0_x64.exe)
- [下载 ZIP（含使用说明与许可证）](https://github.com/treetree-sha/WordMetadataEditor/releases/download/v1.3.0/OfficeMetadataEditor_v1.3.0_x64.zip)
- [下载 SHA-256 校验文件](https://github.com/treetree-sha/WordMetadataEditor/releases/download/v1.3.0/OfficeMetadataEditor_v1.3.0_x64.SHA256.txt)

---

## 🌟 核心功能特性

1. **📄 单文件精细编辑 (Single File Editor)**
   * **拖拽支持**：拖拽支持的 Word、Excel 或 PowerPoint 文件至窗口即可加载文件属性。
   * **核心属性修改**：
     * **作者 (Creator)** & **最后修改者 (Last Modified By)**
     * **创建时间 (Created Time)** & **修改时间 (Modified Time)**（提供“一键设为当前时间”）
     * **总编辑时间 (Total Editing Time)**（单位：分钟，如设为 `120` 表示 2 小时）
     * **修订版本号 (Revision)**
     * 公司名称 (Company) & 文档标题 (Title)
   * **一键脱敏**：快速清空作者、修改人、公司名称，并将总编辑时间归零。

2. **📂 批量处理与脱敏 (Batch Processing)**
   * 支持混合选择多个 Word、Excel、PowerPoint 文件或扫描整个文件夹（递归子目录）。
   * 采用后台多线程处理（UI 界面顺滑不卡顿），支持实时进度条显示与表格状态更新。
   * 支持统一批量设置作者、修改人、编辑时间或一键清理隐私。

3. **⏰ Windows 文件系统时间同步联动 (NTFS Sync)**
   * 勾选“同步更新系统文件时间”后，软件在修改 Office 文件属性的同时，通过 Win32 API 自动修改 Windows NTFS 文件系统的 **【创建时间 (ctime)】** 与 **【修改时间 (mtime)】**。

4. **🎨 Windows 11 Fluent Design 高颜值界面**
   * 支持 **暗黑模式 (Dark Mode)** 和 **浅色模式 (Light Mode)** 无缝切换。
   * 支持 **简体中文 / English 界面语言无缝切换**。

5. **👤 作者信息与 GitHub 联动**
   * 设置页面内置作者 GitHub 个人主页与项目开源仓库链接。

---

## 📂 项目结构

```
WordMetadataEditor/
├── metadata_engine.py      # 统一属性引擎及 Win32 时间同步
├── ole_metadata.py         # Windows 原生 OLE 属性读写和旧格式检查
├── metadata_errors.py      # 文件/加密错误类型
├── main.py                 # 主程序入口 (FluentWindow 导航框架)
├── run.bat                 # Windows 一键启动脚本
├── tests/                  # Office 属性、完整性和界面回归测试
├── README.md               # 中文说明文档
├── README_EN.md            # English Documentation
└── ui/                     # UI 视图模块
    ├── __init__.py
    ├── i18n.py                   # 国际化多语言引擎 (Simplified Chinese & English)
    ├── file_filters.py           # 统一 Office 文件选择过滤器
    ├── single_file_interface.py  # 单文件精修页面
    ├── batch_interface.py        # 批量处理与脱敏页面
    └── settings_interface.py     # 设置与关于页面 (包含语言切换与 GitHub 作者链接)
```

---

## 🚀 启动与打包指南

### 1. 本地启动运行
- **方法一**：在 Windows 资源管理器中双击 **`run.bat`**。
- **方法二**：命令行启动
使用 Python 环境运行 `main.py`：

```bash
python main.py
```

### 2. 验证源码

安装测试依赖 `pytest`、`python-docx`、`openpyxl`、`python-pptx` 后运行（应用运行依赖仍为 `PySide6` 与 `PySide6-Fluent-Widgets`）：

```powershell
python -m pytest tests -q
```

测试覆盖十种格式的属性及容器行为、非属性内容一致性、缺失属性补建、写入失败保护、加密拒绝、Windows 时间同步和 Qt 混合批量处理。Windows OLE 测试使用 `pywin32` 作为独立样本生成/校验依赖；正式软件的旧格式后端仅使用 `ctypes`。真实 Office 样本与独立打开检查见 [测试说明](tests/README.md)。宏流测试使用合成载荷，不执行 VBA。

### 3. 打包为 Windows 单文件 EXE
使用 PyInstaller 进行打包：
```bash
pyinstaller --noconsole --onefile --icon="assets/app_icon.ico" --add-data "assets;assets" --name "OfficeMetadataEditor_v1.3.0_x64" --exclude-module PyQt5 --clean main.py
```

---

## 📄 开源许可证

本项目基于 **[MIT License](LICENSE)** 开源。

Developed with ❤️ by **[treetree-sha](https://github.com/treetree-sha)**.
