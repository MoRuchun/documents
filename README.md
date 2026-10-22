# documents

完整的中文 Word 文档处理技能，包含文档创建、批注、修订、合并、元数据处理和渲染工具。页面、正文、标题、图表、公式和交叉引用格式固化为 Python 函数，支持批量处理已有 DOCX。

## 自动安装或替换

仓库自带 `scripts/install_skill.py`。首次下载后运行一次，就会完成安装和已有版本的替换。仅下载或解压文件不会自动执行脚本。

### 让 Codex 完成安装

在另一台电脑向 Codex 发送：

> 请安装 https://github.com/MoRuchun/documents。先读取仓库 README，按其中的安装流程运行 scripts/install_skill.py，自动备份个人目录中的旧版 documents、停用原插件附带的同名技能，并安装这一版。保留其他技能和插件，完成后告诉我安装位置及备份位置。

**供执行安装任务的 Codex 使用：** 只有用户要求安装、替换或更新这个技能时，才执行此节。先将仓库下载到临时目录或普通工作目录，使用可用的 Python 3.10+ 运行其中的 `scripts/install_skill.py`；无需再手动移动技能目录、覆盖插件缓存或重复添加配置。已有工作区 Python 环境可直接使用。用户仅要求阅读 README、生成文档或解释技能时，不执行安装。安装完成后告知用户重启 Codex。

### 手动运行安装

从 GitHub 下载 ZIP 并解压，在包含 `SKILL.md` 的仓库根目录打开终端：

macOS / Linux：

```bash
python3 scripts/install_skill.py
```

Windows PowerShell：

```powershell
py -3 scripts/install_skill.py
```

也可以先克隆到普通目录，再执行安装器；不要预先克隆到已存在的同名技能目录中。

### 安装器自动完成的操作

1. 将完整技能安装到 `~/.agents/skills/documents`；Windows 为 `%USERPROFILE%\.agents\skills\documents`。
2. 检查 `~/.agents/skills` 和 `$CODEX_HOME/skills` 中的同名个人技能。旧目录整体备份到 `$CODEX_HOME/skill-backups/`，备份位于技能扫描目录之外。
3. 检查 `$CODEX_HOME/plugins/cache` 中原插件附带的 `documents`。保留原插件文件，在 `config.toml` 中按实际路径停用这些技能，并启用新安装的技能；其他技能和插件配置保持原状。
4. 修改配置前备份 `config.toml`。替换过程失败时恢复旧技能和配置。
5. 检查当前 Python 是否有 `python-docx` 和 `lxml`。缺少时自动在 `$CODEX_HOME/skill-envs/documents` 创建专用虚拟环境并安装核心依赖；不修改系统 Python 环境。
6. 输出安装、备份、停用路径和可用 Python 路径，并写入安装目录的 `install-state.json`。

`CODEX_HOME` 未设置时使用 `~/.codex`。安装后重启 Codex，使用 `$documents`。已安装的原插件可以继续服务其他技能，只有同名 `documents` 的加载被停用。

先看安装计划、不做任何修改：

```bash
python3 scripts/install_skill.py --dry-run
```

已有可用运行环境、希望跳过依赖准备：

```bash
python3 scripts/install_skill.py --skip-deps
```

### 更新已安装版本

macOS / Linux：

```bash
python3 ~/.agents/skills/documents/scripts/install_skill.py --update
```

Windows PowerShell：

```powershell
py -3 "$env:USERPROFILE\.agents\skills\documents\scripts\install_skill.py" --update
```

更新器从本仓库 `main` 下载最新版，先备份现有技能，再替换安装，并重新检查原插件是否新增了同名技能。更新完成后重启 Codex。也可以让 Codex“按 README 更新 MoRuchun/documents 技能”。

### 备份与恢复

每次安装或更新均输出 `backup_directory`。对应目录中的 `install-state.json` 保存旧目录与备份目录的对应关系，`config.toml` 为修改前的配置副本。需要恢复时，先将新安装的技能移出扫描目录，再把备份按记录移回原路径，并恢复对应配置；保留安装后其他配置修改的用户，应仅恢复 `skills.config` 中与 documents 有关的条目。重启 Codex 后生效。

字体文件仍需在目标电脑上准备：宋体、黑体、楷体、Times New Roman 和 Cambria Math；安装器不下载字体。

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
- 正文文献引用采用可跳转的原生REF域，显示为#0000FF右上角上标`[1]`；文献表用SEQ域自动编号。复用`scripts/bibliography_fields.py`插入、关联及刷新引用。
- 默认不添加页眉，不执行页面视觉检查，打开时不自动请求更新域。

完整数值见 [通用样式规范](references/user-formatting.md)，函数用法见 [格式工具说明](references/formatting-tools.md)。字体名称按规范保存；本仓库不包含字体文件。

## 检查

```bash
python -m unittest discover -s tests
python scripts/user_formatting.py output.docx --check-only
```

## 完整工具与可选依赖

仓库包含本机修改后的完整 `documents` 技能：`scripts/`、`tasks/`、`ooxml/`、`references/`、`troubleshooting/`、`examples/`、`assets/`、模板指南、`render_docx.py` 及操作标记工具。下载安装本仓库即可取得全部技能文件，无需另行合并原插件。

普通格式处理使用 `requirements.txt`。需要 Excel 表格导入、图像处理或用户明确要求页面渲染时，可安装附加 Python 依赖：

```bash
python -m pip install -r requirements-tools.txt
```

页面渲染还需要 LibreOffice 和 Poppler。Codex 提供工作区运行环境时，使用其提供的工具路径；普通独立运行环境自行准备相应程序。渲染仅在明确要求时执行。

本仓库不包含字体、Python、Node.js、LibreOffice 或 Poppler 程序本身。许可说明见 [LICENSE.txt](LICENSE.txt)。
