<p align="center">
  <strong>English</strong> | <a href="README.zh-CN.md">简体中文</a>
</p>

# office-skill

**Token-compact, guardrail-first `.xlsx`, `.docx`, and `.pptx` skill for AI coding agents.**

Cross-platform (**Linux / WSL / macOS / Windows**). Built for beginners who just want their agent to read and write Office files reliably without blowing up context windows, shipping broken `#REF!` spreadsheets, or corrupting Word layouts.

---

## One-Line Install (Let Your Agent Do It)

Copy and paste this sentence to your AI agent (Antigravity, Claude Code, Codex, Cursor, etc.):

> Read https://raw.githubusercontent.com/26kyo/office-skill/main/AGENT_INSTALL.md and install the `office` skill for me, then run its self-check and report the result.

---

## Why This Skill Exists

Agents routinely fail on Office files in three predictable ways:
1. **Context bloat** — dumping raw XML or verbose JSON for a 500-row sheet burns thousands of tokens.
2. **Silent formula bugs** — `openpyxl` writes formulas without evaluating them, so `#REF!` / `#DIV/0!` / `#VALUE!` slip through unnoticed.
3. **Broken `.docx` layouts** — rewriting existing Word files through `python-docx` drops unsupported OOXML features or collapses table columns across Word / LibreOffice / Google Docs.

### What `office-skill` Provides

| Script | Purpose |
| :--- | :--- |
| `office/scripts/office_read.py` | Read-only, token-compact inspector: Excel TSV + `# formulas` (`formula => cached value`), Word outline/content, PPTX overview/slides (including grouped shapes) + optional PNG rendering. |
| `office/scripts/xlsx_recalc.py` | Headless LibreOffice recalculation with isolated `UserInstallation` profile + 7-error Excel scanner (`#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, `#NUM!`, `#NULL!`) + non-zero exit gate + `--backup` / `--no-write`. |
| `office/scripts/docx_xml.py` | Safe OOXML `unpack` / `validate` / `pack` with zip-slip protection (`is_relative_to`) and byte-preserving default unpack for surgical XML edits. |
| `office/scripts/doctor.py` | Zero-dependency self-check reporting Python, `uv`, modules, LibreOffice, and `pdftoppm` status with OS-specific install commands. |

---

## Manual Install (Optional)

1. Copy the `office/` directory into your agent's skills folder (e.g. `~/.gemini/config/skills/office`, `~/.claude/skills/office`, `~/.codex/skills/office`, or `~/.cursor/skills/office`).
2. **With [`uv`](https://docs.astral.sh/uv/) (recommended)**: nothing else to install — all scripts declare PEP 723 inline dependencies and run directly via `uv run`.
3. **Without `uv`**:
   ```bash
   python3 -m venv office/.venv
   office/.venv/bin/pip install -r office/requirements.txt
   ```
4. **Optional system tools** (for Excel formula recalculation & slide/doc PNG preview):
   - Ubuntu / Debian / WSL: `sudo apt update && sudo apt install -y libreoffice poppler-utils`
   - macOS: `brew install --cask libreoffice && brew install poppler`
   - Windows: `winget install TheDocumentFoundation.LibreOffice && winget install oschwartz10612.Poppler`
5. Verify:
   ```bash
   python3 office/scripts/doctor.py
   ```

---

## Running Tests

```bash
python3 -m unittest discover -s tests -v
```

Tests dynamically generate minimal `.xlsx`, `.docx`, and `.pptx` fixtures in a temporary directory — zero binary fixtures committed to git.

---

## Credits & License

- **Independent clean-room implementation** licensed under the [MIT License](LICENSE).
- Inspired by document workflow guardrails popularized in [`anthropics/skills`](https://github.com/anthropics/skills) (headless LibreOffice formula recalculation gating, OOXML unpack/validate/pack editing, and dual-width Word tables). No code or text was copied from that repository.
- Powered by open-source tools: `openpyxl` (MIT), `python-docx` (MIT), `python-pptx` (MIT), `lxml` (BSD-3-Clause), `pandas` (BSD-3-Clause), `docxtpl` (LGPL-2.1), `LibreOffice` (MPL-2.0), and `Poppler` (GPL, invoked as an external CLI).
