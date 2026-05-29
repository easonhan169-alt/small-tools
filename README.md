# Excel 文档对比与替换工具

一个基于 Streamlit 的 Excel 文档对比软件原型，用于上传两个 Excel 工作簿或 CSV 文件，输出详细差异，并提供一键生成替换后的工作簿。

## 功能

- 对比两个 `.xlsx` / `.xlsm` / `.xltx` / `.xltm` / `.csv` 文件。
- 输出工作表结构差异：新增/缺失工作表、行列尺寸变化。
- 输出单元格数据差异：工作表、单元格坐标、左侧值、右侧值、差异类型。
- 输出字符级差异：对文本单元格使用 `difflib` 生成 HTML 高亮片段。
- 支持一键替换：下载“左侧文件替换为右侧内容”或“右侧文件替换为左侧内容”的新 Excel/CSV 文件。
- 支持导出差异明细 CSV。

## 快速开始

### Windows PowerShell（推荐）

如果你在 Windows PowerShell 里运行，请不要使用 `source .venv/bin/activate`，这是 Linux/macOS 的命令。推荐直接用虚拟环境里的 `python.exe` 启动，这样不会受 PATH 环境变量影响：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run src/excel_diff_app.py
```

如果你已经用 `pip install -r requirements.txt` 安装成功，但运行 `streamlit` 提示“无法将 streamlit 项识别为 cmdlet”，说明 `streamlit.exe` 所在目录没有加入 PATH。可以直接改用：

```powershell
python -m streamlit run src/excel_diff_app.py
```

如果你想激活虚拟环境，PowerShell 命令是：

```powershell
.\.venv\Scripts\Activate.ps1
```

如果激活时提示脚本执行策略限制，可以临时允许当前 PowerShell 会话执行本地脚本：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

### macOS / Linux

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run src/excel_diff_app.py
```

## 使用方式

1. 在页面左侧上传“左侧/原始文档（Excel/CSV）”。
2. 上传“右侧/目标文档（Excel/CSV）”。
3. 查看结构差异、数据差异和字符级差异。
4. 点击下载按钮生成替换后的 Excel/CSV 文件。

> 为了避免意外覆盖原文件，本工具不会直接修改上传文件，而是生成新的可下载文件。
