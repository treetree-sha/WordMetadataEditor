# 属性引擎与界面验证

正式软件依赖 `PySide6`、`PySide6-Fluent-Widgets`。新增的 Windows OLE 后端仅用标准库 `ctypes`，不依赖 Office 或 `pywin32`。

测试额外依赖：`pytest`、`python-docx`、`openpyxl`、`python-pptx`。Windows OLE 测试使用 `pywin32` 独立生成复合文件和读取数据流；这是测试依赖。

在项目根目录运行：

```powershell
python -m pytest tests -q
```

如运行环境限制系统临时目录，可将临时数据放入项目已忽略的测试目录（选用新的子目录）：

```powershell
python -m pytest tests -q --basetemp=test_documents/pytest_local
```

常规测试覆盖：

- 已有现代 Office 格式及 XLSB 的属性读写、内容类型声明、部件内容保留。
- DOC/XLS/PPT 的格式与加密检查、中文和长备注、时长和时间、缺失属性补建。
- OLE 未知属性、管理者信息、同流中的自定义属性及嵌套 VBA/非属性流保留。
- 写入失败与扩展名错误时的原文件保护。
- 单文件表单、拖拽、加密提示、中英文切换及真实 Qt 后台线程混合处理。

OLE 单元测试的正文流和 VBA 流为合成载荷，用于验证容器及属性行为，不作为可被 Office 打开的完整文件。没有真实 XLSB 样本时，相应单元测试使用包含不透明二进制载荷的包级样本。

## 可选：真实 Office 文件与独立打开验证

开发机安装 Microsoft Word、Excel、PowerPoint 和 `pywin32` 后，可以运行以下开发辅助工具，在独立后台 Office 实例中生成测试文件并验证。生成文件位于 Git 已忽略的 `test_documents` 下。

```powershell
python tools/native_office_check.py --generate
```

该命令要求目标目录中尚无 `sample.doc`、`sample.xls`、`sample.xlsb`、`sample.ppt`，以避免覆盖原始样本。已生成样本后直接复验：

```powershell
python tools/native_office_check.py
python -m pytest tests -q
```

测试会复制原始文件，在副本上修改属性，再通过相应 Office 应用以只读方式重新打开，核对中文属性、Word 正文、Excel 公式/结果/格式和 PowerPoint 幻灯片。原始文件按 SHA-256 校验保持不变。检查结果保存为 `office-check-*/native_check.json`。

真实样本也可通过 `OFFICE_NATIVE_FIXTURES` 环境变量指定目录；开发辅助工具使用 `--fixtures` 参数指定目录。缺少真实样本时，四项可选真实样本测试会跳过。

实测 Word 打开 `.doc` 时可能从文档内部读取总编辑时长，而非文件摘要中的值；工具只修改摘要属性，不修改 Word 正文流中的内部统计。辅助检查会记录摘要值与应用显示值，界面和说明也明确标注这一限制。
