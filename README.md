# Word 文档属性高级编辑器 (Word Metadata Editor)

![Windows 11 Fluent Design](https://img.shields.io/badge/Style-Win11%20Fluent%20Design-0078D4?style=flat-square)
![Python 3.9+](https://img.shields.io/badge/Python-3.9%2B-blue?style=flat-square)
![PySide6](https://img.shields.io/badge/GUI-PySide6%20%2B%20FluentWidgets-green?style=flat-square)

一款基于 **PySide6** 与 **PyQt-Fluent-Widgets** 打造的高颜值 Word (`.docx`) 文档元数据及文件系统属性修改工具。

---

## 🌟 核心功能

1. **📄 单文件精细编辑**
   * **拖拽支持**：拖拽任意 `.docx` 文件至窗口即可自动解析并加载所有元数据。
   * **核心属性修改**：
     * **作者 (Creator)** & **最后修改者 (Last Modified By)**
     * **创建时间 (Created Time)** & **修改时间 (Modified Time)**（包含“一键设为当前时间”方便快捷输入）
     * **总编辑时间 (Total Editing Time)**（单位：分钟，如设为 `120` 表示 2 小时）
     * **修订版本号 (Revision)**
     * 标题、主题、公司名称等
   * **一键脱敏**：快速清空作者、修改人、公司名称，并将编辑时间归零。

2. **📂 批量处理与脱敏**
   * 支持同时选择多个 Word 文档或扫描整个文件夹（递归包含子目录）。
   * 采用后台多线程处理（不阻塞 GUI 界面），支持实时进度条显示与表格状态更新。
   * 支持统一修改批量文件的作者、编辑时间或一键清理隐私。

3. **⏰ Windows 文件系统时间同步联动**
   * 勾选“同步更新系统文件时间”后，软件在修改 `.docx` 内部 XML 的同时，通过 Win32 API 自动改变 Windows NTFS 文件系统的 **【创建时间 (ctime)】** 与 **【修改时间 (mtime)】**。

4. **🎨 Windows 11 Fluent Design 高颜值界面**
   * 支持 **暗黑模式 (Dark Mode)** 和浅色模式无缝切换。
   * 毛玻璃/亚克力感侧边导航栏、平滑交互动画与优雅的卡片布局。

---

## 📂 项目结构

```
wordediter/
├── metadata_engine.py      # 元数据核心引擎 (解析/修改 core.xml、app.xml 及 Win32 API 时间同步)
├── main.py                 # 主程序入口 (FluentWindow 导航框架)
├── run.bat                 # Windows 一键启动脚本
├── test_document.docx      # 默认示例测试文档
├── README.md               # 项目说明文档
└── ui/                     # UI 视图模块
    ├── __init__.py
    ├── single_file_interface.py  # 单文件精修页面 (支持拖拽)
    ├── batch_interface.py        # 批量处理与脱敏页面
    └── settings_interface.py     # 设置与主题切换页面
```

---

## 🚀 快速启动

### 方法一：双击脚本启动（推荐）
在 Windows 资源管理器中直接双击运行 **`run.bat`** 即可。

### 方法二：命令行启动
使用 Python 环境运行 `main.py`：

```bash
python main.py
```

*如果使用 Anaconda 环境，可直接执行：*
```bash
D:\Programs\Anaconda\python.exe main.py
```

---

## 📦 依赖库说明

运行本程序需要以下 Python 依赖库（已自动安装）：
- `PySide6` & `PySide6-Fluent-Widgets`：用于构建 Windows 11 风格现代 GUI
- `python-docx` & `lxml`：用于解析及辅助处理 Word XML 结构
- `pywin32`：用于调用 Windows Win32 API 修改文件系统时间戳

---

## 🔒 隐私与安全性

* 本工具为本地离线软件，不会上传任何文档或元数据到网络。
* 修改操作直接作用于本地 `.docx` 文件容器内的 `docProps/core.xml` 与 `docProps/app.xml`。
