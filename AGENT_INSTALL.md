# Agent Installation Guide (`AGENT_INSTALL.md`)

> Instructions for AI coding agents (Antigravity, Claude Code, Codex, Cursor, Windsurf, etc.) to install and verify the `office` skill automatically on behalf of the user.

## Step 1: Locate Your Agent's Skills Directory

Detect the active agent environment and pick the matching skills root (ask the user only if ambiguous):

| Agent / Tool | Typical Global Skills Directory |
| :--- | :--- |
| Antigravity | `~/.gemini/config/skills/` |
| Claude Code | `~/.claude/skills/` |
| Codex | `~/.codex/skills/` |
| Cursor | `~/.cursor/skills/` |
| Project-local (any agent) | `<project>/.agents/skills/` or `<project>/.claude/skills/` |

Target directory for this skill: `<SKILLS_DIR>/office`.

## Step 2: Fetch & Copy the `office/` Folder

Clone or download this repository into a temporary folder and copy `office/` into `<SKILLS_DIR>/office`:

```bash
git clone --depth 1 https://github.com/26kyo/office-skill.git /tmp/office-skill-install
mkdir -p "<SKILLS_DIR>"
cp -r /tmp/office-skill-install/office "<SKILLS_DIR>/office"
rm -rf /tmp/office-skill-install
```

*(On Windows PowerShell, use `Copy-Item -Recurse` into the user's home skills path.)*

If `<SKILLS_DIR>/office` already exists, confirm with the user before overwriting.

## Step 3: Set Up the Python Runtime

1. Check if `uv` is available (`uv --version`):
   - **If `uv` is installed**: no virtual environment is needed! Every script in `office/scripts/` contains PEP 723 inline dependencies and runs directly via `uv run`.
   - **If `uv` is NOT installed**: create a self-contained venv inside the skill folder:
     ```bash
     python3 -m venv "<SKILLS_DIR>/office/.venv"
     "<SKILLS_DIR>/office/.venv/bin/pip" install -r "<SKILLS_DIR>/office/requirements.txt"
     ```
     *(Windows: `"<SKILLS_DIR>\office\.venv\Scripts\pip.exe" install -r "<SKILLS_DIR>\office\requirements.txt"`)*

## Step 4: Run Environment Self-Check

Run `doctor.py` (uses Python standard library only):

```bash
python3 "<SKILLS_DIR>/office/scripts/doctor.py"
```

- If `Core read/write ready: True`, core `.xlsx` / `.docx` / `.pptx` read and write works immediately.
- If `LibreOffice` or `pdftoppm` is missing:
  - These are optional for basic reading/writing, but **required** for Excel formula recalculation (`xlsx_recalc.py`) and PDF/PNG visual layout checks (`office_read.py pptx --render`).
  - Show the OS-specific install command printed by `doctor.py` and ask the user before running any system package manager (`apt` / `brew` / `winget`).
  - **WSL note**: install Linux-native `libreoffice` and `poppler-utils` inside WSL; do not invoke Windows `.exe` binaries from WSL.

## Step 5: Report Completion

Confirm to the user:
1. Where the skill was installed (`<SKILLS_DIR>/office`).
2. Which runtime mode is active (`uv run` or `.venv`).
3. Whether `LibreOffice` and `pdftoppm` are ready (or the one-line command to add them later).
