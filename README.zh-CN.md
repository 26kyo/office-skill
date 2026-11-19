<p align="center">
  <a href="README.md">English</a> | <strong>简体中文</strong>
</p>

# office-skill

**面向 AI 编程 Agent 的省 Token、强校验 `.xlsx` / `.docx` / `.pptx` 读写 Skill。**

跨平台支持 **Linux / WSL / macOS / Windows**。专为想让 Agent 无脑、可靠读写 Office 文件的新手设计：不炸上下文、不产出带 `#REF!` 的坏表格、不破坏已有 Word 排版。

---

## 一句话安装（直接让你的 Agent 装）

把下面这句话复制发给你的 AI Agent（Antigravity、Claude Code、Codex、Cursor 等）：

> 请读取 https://raw.githubusercontent.com/26kyo/office-skill/main/AGENT_INSTALL.md ，帮我安装 `office` skill，装完运行自检并汇报结果。

---

## 为什么需要这个 Skill

让 Agent 直接操作 Office 文件，最常见的三个翻车点：
1. **上下文爆炸** — 把几百行的表格转成冗长 XML/JSON 喂给模型，瞬间烧光 Token。
2. **公式静默报错** — `openpyxl` 写完公式不会自动计算，`#REF!` / `#DIV/0!` / `#VALUE!` 混进交付文件里毫无察觉。
3. **Word 排版崩坏** — 用 `python-docx` 重写已有文档容易丢批注/修订等 OOXML 特性，表格列宽在 Word / LibreOffice / Google Docs 之间也常塌陷。

### 内置脚本一览

| 脚本 | 作用 |
| :--- | :--- |
| `office/scripts/office_read.py` | 只读、极简 Token 检视器：Excel 输出紧凑 TSV + `# formulas`（`公式 => 缓存值`）、Word 大纲/正文、PPT 概览/单页（支持递归组合形状）+ 可选 PNG 渲染。 |
| `office/scripts/xlsx_recalc.py` | 使用独立 `UserInstallation` profile 调 headless LibreOffice 重算公式，扫描 7 类 Excel 错误（`#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A`、`#NUM!`、`#NULL!`），有错即非零退出，支持 `--backup` / `--no-write`。 |
| `office/scripts/docx_xml.py` | 安全的 OOXML 拆包 / XML 语法校验 / 打包工具，带 zip-slip 路径防护（`is_relative_to`），默认无损保留原始字节以便精准改 XML。 |
| `office/scripts/doctor.py` | 零第三方依赖的环境自检工具，一键检查 Python、`uv`、依赖库、LibreOffice、pdftoppm，并按当前系统给出缺失项安装命令。 |

---

## 手动安装（可选）

1. 将 `office/` 目录复制到你所用 Agent 的 skills 目录（如 `~/.gemini/config/skills/office`、`~/.claude/skills/office`、`~/.codex/skills/office` 或 `~/.cursor/skills/office`）。
2. **已装 [`uv`](https://docs.astral.sh/uv/)（推荐）**：无需任何配置，所有脚本自带 PEP 723 内联依赖，直接 `uv run` 运行。
3. **未装 `uv`**：
   ```bash
   python3 -m venv office/.venv
   office/.venv/bin/pip install -r office/requirements.txt
   ```
4. **可选系统依赖**（用于 Excel 公式重算与文档/PPT 视觉检查）：
   - Ubuntu / Debian / WSL：`sudo apt update && sudo apt install -y libreoffice poppler-utils`
   - macOS：`brew install --cask libreoffice && brew install poppler`
   - Windows：`winget install TheDocumentFoundation.LibreOffice && winget install oschwartz10612.Poppler`
5. 自检：
   ```bash
   python3 office/scripts/doctor.py
   ```

---

## 运行测试

```bash
python3 -m unittest discover -s tests -v
```

测试会在临时目录动态生成最小 `.xlsx` / `.docx` / `.pptx` 样本，仓库内零二进制垃圾。

---

## 致谢与许可

- 本项目为**独立净室实现（clean-room implementation）**，采用 [MIT License](LICENSE) 开源。
- 工作流护栏思路受 [`anthropics/skills`](https://github.com/anthropics/skills) 启发（headless LibreOffice 公式重算门禁、OOXML 拆包校验编辑、Word 双宽度表格）。未复制该仓库的任何代码或文本。
- 底层依赖开源生态：`openpyxl` (MIT)、`python-docx` (MIT)、`python-pptx` (MIT)、`lxml` (BSD-3-Clause)、`pandas` (BSD-3-Clause)、`docxtpl` (LGPL-2.1)、`LibreOffice` (MPL-2.0)、`Poppler` (GPL，仅作为外部命令行调用)。
