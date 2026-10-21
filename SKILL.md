---
name: documents
description: Create and format DOCX files with reusable tools for Chinese document styles, multilevel headings, tables, native equations, captions, and clickable cross-references.
---

# Documents

创建或统一 DOCX 格式时，读取 [通用样式规范](references/user-formatting.md)。该规范用于已有内容的样式设置；文档内容由用户任务决定。用户另行指定的格式或模板优先，局部内容修改保留原有格式。

使用 [格式工具说明](references/formatting-tools.md) 中的 `scripts/user_formatting.py`。新建文档导入其中的样式及元素函数；批量处理已有文档使用：

```bash
python scripts/user_formatting.py input.docx --out output.docx
```

- 已有文档采用其他样式名称时，用 `--style-map mapping.json` 显式映射到规范中的样式；不要根据外观猜测内容角色。
- 复用工具中的编号、表格、题注、原生公式和交叉引用处理逻辑，避免每次重新实现。
- 新建文档完成内容后执行 `apply_document(doc)` 和 `check_document(doc)` 再保存。
- 行内、行间数学内容均使用原生 Word 公式对象 OMML；数学字体 Cambria Math。
- 默认不渲染页面、不生成页面截图、不做视觉检查；用户明确要求时再执行。
- 打开时自动更新域默认关闭；保留可更新的域及其当前显示结果。

## 运行环境

优先使用当前环境已提供的 Python 及依赖。若 Codex 提供工作区依赖加载工具，先解析其运行环境；普通 Python 环境的依赖列在 `requirements.txt`，按需安装。脚本自身不安装软件，不依赖特定电脑的路径。

## 工具入口

| 用途 | 函数 |
| --- | --- |
| 创建带完整样式的文档 | `new_document()` |
| 批量应用全部格式 | `apply_document(doc, style_map=None)` |
| 一级、二级、三级标题 | `add_heading(doc, text, level)` |
| 图题、表题与续表题 | `add_caption(doc, kind, title, bookmark, chapter, index, continued=False)` |
| 三线表 | `add_table(doc, headers, rows, widths_cm=None)` |
| 可点击交叉引用 | `add_cross_reference(paragraph, bookmark, cached_number, prefix="")` |
| 原生行内公式 | `add_inline_math(paragraph, elements)` |
| 带自动编号的行间公式 | `add_numbered_equation(doc, elements, bookmark, chapter, index)` |
| 文档结构检查 | `check_document(doc)` 或命令行 `--check-only` |

续表位置、目录和分节由文档构建时组织；工具批量设置样式，不依据未渲染页面猜测分页。交付用户要求的文件。
