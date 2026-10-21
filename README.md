# documents

可复用的中文 Word 文档格式技能。将页面、正文、标题、图表、公式和交叉引用格式固化为 Python 函数，支持批量处理已有 DOCX。

## 在另一台电脑上使用

将本仓库下载或克隆到技能目录，文件夹名称保留为 `documents`。macOS / Linux 的默认路径示例：

```bash
git clone https://github.com/MoRuchun/documents.git ~/.codex/skills/documents
```

Windows 可克隆到 `%USERPROFILE%\.codex\skills\documents`。如果该目录已有同名技能，先将原目录移到其他位置，再放入本技能。

开始新会话后使用 `$documents`。也可以直接运行脚本，无需通过助手操作。

## 直接批量处理文档

脚本需要 Python 3.10 或更新版本。已有环境提供依赖时直接使用；需要自行准备环境时，在仓库目录中创建虚拟环境：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python scripts/user_formatting.py input.docx --out output.docx
```

Windows 的虚拟环境命令使用 `.venv\Scripts\python.exe`。输入与输出可指定为同一路径。

原文档使用不同样式名称时，可提供映射：

```json
{"二级标题":"Heading 2","文献条目":"参考文献"}
```

```bash
python scripts/user_formatting.py input.docx --out output.docx --style-map mapping.json
```

## 当前样式

- A4 纵向；上、下 2.5 cm，左 3.0 cm，右 2.5 cm。
- 正文宋体与 Times New Roman，小四，固定 22 磅行距，首行缩进 2 字符。
- 黑体三级标题，自动编号为 1、1.1、1.1.1。
- 三线表，表内五号、左对齐、单倍行距，表头不加粗。
- 原生可编辑 Word 公式，Cambria Math；图表、公式按章编号。
- 可点击交叉引用，全部引用文字统一为 #0000FF。
- 参考文献左缩进 0，悬挂缩进 2 字符。
- 默认不添加页眉，不执行页面视觉检查，打开时不自动请求更新域。

完整数值见 [通用样式规范](references/user-formatting.md)，函数用法见 [格式工具说明](references/formatting-tools.md)。字体名称按规范保存；本仓库不包含字体文件。

## 检查

```bash
python -m unittest discover -s tests
python scripts/user_formatting.py output.docx --check-only
```

仓库包含本次定制的格式规范与工具，不包含原 Documents 插件的附带文件。
