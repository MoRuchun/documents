# 通用样式工具

使用 `scripts/user_formatting.py` 实现 `references/user-formatting.md` 中的已确认格式。保留说明文件作为参数规范；由脚本执行样式、段落、编号、表格、公式字体与交叉引用的批量设置，不在每次任务中重新编写这些格式代码。

## 批量处理已有文档

使用依赖加载器返回的 Python 运行：

```bash
python scripts/user_formatting.py input.docx --out output.docx
```

- 根据已有内容的样式角色处理，不增加章节或其他内容。
- 自动处理内置 `Heading 1`、`Heading 2`、`Heading 3`，以及格式规范中的命名样式。已有文档采用其他样式名称时，用 JSON 对象映射名称，例如 `{"二级标题":"Heading 2","文献条目":"参考文献"}`，通过 `--style-map mapping.json` 传入。
- 自动识别所有表格单元格并设置表内样式。图题、表题、文献等语义由样式或显式映射确定，不凭外观猜测。
- 保留正文、图片、原生数学对象、书签及域，更新指定格式。
- 完成后检查标题的实际编号关联、整段交叉引用颜色、每个单元格的单倍行距和参考文献缩进。仅检查已有文件可使用 `--check-only`，不会修改文件。

## 新建文档

把技能的 `scripts` 目录加入模块路径后调用：

```python
from user_formatting import (
    new_document, add_heading, add_caption, add_table,
    add_cross_reference, apply_document, check_document,
)

doc = new_document()
add_heading(doc, "观测方法", level=1)
add_heading(doc, "记录方式", level=2)
add_heading(doc, "单位与时间", level=3)
doc.add_paragraph("这里填写正文。", "Normal")
add_caption(doc, "表", "地块记录", "PlotRecords", chapter=1, index=1)
add_table(doc, ["地块", "面积（ha）"], [["林荫地块", "0.85"]], [8, 7.5])
p = doc.add_paragraph("记录见", "Normal")
add_cross_reference(p, "PlotRecords", "1-1", prefix="表 ")
apply_document(doc)
check_document(doc)
doc.save("output.docx")
```

`add_caption(..., continued=True)` 创建重复原表号和表题的续表题。`add_inline_math(paragraph, elements)` 接收 OMML 元素；`add_numbered_equation(doc, elements, bookmark, chapter, index)` 设置行间公式的位置与自动编号。公式结构使用原生 OMML，字体由批量工具统一设置。

## 文献引用原生域

用`scripts/bibliography_fields.py`处理文献引用。已有文献表的条目先设为`参考文献`样式，保留`[1] 条目文字`的数字标记；再调用`link_numeric_citations(doc)`，将正文对应标记转换为可点击的蓝色上标`REF`域，将条目编号转换为带书签的`SEQ BibliographyItem`域。函数保留其他文字格式，可处理跨文字片段的引用，重复调用保持既有链接。多篇引用如`[1,2]`逐个关联相应文献。

```python
from bibliography_fields import (
    add_bibliography_entry, add_citation, link_numeric_citations,
    refresh_bibliography_cache, check_bibliography_fields,
)

# 新文档：key使用稳定的ASCII文献标识。
entry = add_bibliography_entry(doc, "作者. 题名[J]. 期刊, 年份.", "author2024", 1)
p = doc.add_paragraph("相关研究指出……", "Normal")
add_citation(p, "author2024", 1)

# 已有文档：为编号条目和正文引用建立关联。
link_numeric_citations(doc)
apply_document(doc)
check_document(doc)

# 文献条目顺序变化后，按文献表顺序刷新SEQ与REF缓存。
refresh_bibliography_cache(doc)
check_bibliography_fields(doc)
```

独立转换：`python scripts/bibliography_fields.py input.docx --out output.docx`。`apply_document`会保持已有文献域的蓝色上标，`check_document`同时核对链接目标与上标颜色。`link_numeric_citations`适用于一个连续编号的文献表；多个独立编号范围需分别处理，重复编号会报告歧义。

## 其他格式实现约定

- 标题的段落本身显式写入 `numId` 与 `ilvl`；编号模板同时关联对应标题样式。三级编号分别为 `%1`、`%1.%2`、`%1.%2.%3`，低层级按上层级重置。使用 Word 默认的逐级重启行为，不写入显式 `lvlRestart`；已在 Word 中读取实际列表文本确认子级编号正常显示。
- 可点击交叉引用的完整内容统一设置 #0000FF，包括前缀、数字和连接符；域使用 `CHARFORMAT` 保留更新后的颜色。
- 表内单倍行距明确写入样式及每个单元格段落：`line=240`、`lineRule=auto`，清除局部行距和段落缩进覆盖。
- 参考文献左缩进为 0，悬挂缩进为 2 字符；样式及实际条目同时设置。
- 续表页的位置、目录和页码等内容结构由文档构建时组织；批量格式工具不根据未渲染的页面猜测分页。
- 默认不生成页面图片、不执行视觉检查。

打开文件时不自动请求更新域（`updateFields=false`）；保留域及其显示结果，按需在 Word 中手动更新。
