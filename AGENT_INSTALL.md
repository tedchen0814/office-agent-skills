# Installing office-agent-skills — instructions for AI agents

Follow these steps when the user asks you to install this repo (for example: "install this skill: https://github.com/tedchen0814/office-agent-skills"). Talk to the user in their language (Traditional Chinese if they wrote in Chinese).

## What gets installed

- Skills `html-slide-builder`, `soil-html-deck`, `soil-image-deck`, `soil-teaching-deck`, `draw`, `file-toolkit` (this repo), into the agent's skills folder
- Python packages `openai`, `pillow`, `python-pptx`, `python-docx`, `openpyxl`, `pypdf`, `reportlab`, `pymupdf`, `qrcode[pil]`, `markitdown[pdf,docx,pptx,xlsx]` (`matplotlib` is deliberately not installed; install it only when a chart is needed)
- npm global package `pptxgenjs@3.12.0` (only if Node.js is present)

Existing skills with the same names are moved to `~/.claude/skill-backups/<timestamp>/`, not deleted. Nothing is written outside the skills folder, the backup folder, and the pip / npm package locations. **The installer never asks for, stores, or downloads an OpenAI key or a Firebase project.**

## Steps

1. **Pick the skills folder.** Claude Code → default (`~/.claude/skills`). Codex → add `-Agent codex` / `--agent codex`. If unsure, ask the user.

2. **Check first, change nothing.**
   - Windows (PowerShell):
     ```powershell
     & ([scriptblock]::Create((irm https://raw.githubusercontent.com/tedchen0814/office-agent-skills/main/install.ps1))) -CheckOnly
     ```
   - macOS / Linux:
     ```bash
     curl -fsSL https://raw.githubusercontent.com/tedchen0814/office-agent-skills/main/install.sh | bash -s -- --check-only
     ```

3. **Missing prerequisites?** Git and Python 3.9+ are required; Node.js 22+ and GitHub CLI are optional (see README for what each unlocks). Installing system software needs the user's OK. On Windows, with their OK, rerun with `-InstallPrereqs` (winget). On macOS/Linux give them the `brew` / `apt-get` command the script prints.

4. **Firebase (optional).** Ask once: "Do you want live word-cloud / poll slides? That needs your own Firebase project." If yes, ask them to paste their `firebaseConfig` (Firebase Console → Project settings → Your apps → Web), save it to a temp text file, and add `-FirebaseConfigFile <file>` / `--firebase-config <file>`. Do not reuse anyone else's project. If no, skip: `html-slide-builder` then leaves interactive slides out.

5. **Run the installer** (same command as step 2, without the check flag). Use a timeout of at least 10 minutes.

6. **OpenAI key.** The report ends with whether a key was found. If not, tell the user image generation needs their own key, and show them the two steps from the README (create `~/.openai.env` with `OPENAI_API_KEY=...`; complete Individual verification). Do not ask them to paste the key into the chat; let them create the file themselves.

7. **Report** the "Environment check" section: every skill should show `[ok]`. Then tell the user to **restart the agent** so the new skills load, and give one example:
   ```text
   把這份大綱做成互動 HTML 簡報：<貼上大綱>
   ```

## If something fails

- `git clone` fails with "Filename too long" (Windows): the script already sets `core.longpaths=true`; if it still fails, run `git config --global core.longpaths true` with the user's OK and retry.
- "is not digitally signed" (Windows): group policy blocks scripts on network or cloud drives. Use the `irm` scriptblock one-liner (it does not run a file), or copy the repo to a local disk such as `C:\` and run it there.
- `python` exits with code 49 (Windows): that is the Microsoft Store stub. The script tries `py -3` first; if no real Python exists, install one (step 3).
- `npm install -g` permission error (macOS/Linux): do not use `sudo` without asking; suggest fixing npm's global prefix (`npm config set prefix ~/.npm-global` and adding it to PATH).
- `pip install` says "externally managed environment" (Linux/macOS): the script retries with `--break-system-packages`; prefer a virtual environment if the user has one.
- Needs github.com, raw.githubusercontent.com, registry.npmjs.org and pypi.org.

Do not edit the downloaded skill files during installation.
