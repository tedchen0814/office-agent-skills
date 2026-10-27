#!/usr/bin/env bash
# office-agent-skills one-step installer (macOS / Linux / Git Bash)
#
# Installs html-slide-builder, soil-html-deck, soil-image-deck, soil-teaching-deck, draw and file-toolkit
# into the agent's skills folder, plus the Python / npm packages they need.
# It never installs or stores an OpenAI key.
#
# Usage (from a clone):  bash install.sh [options]
# Usage (no clone):      curl -fsSL https://raw.githubusercontent.com/tedchen0814/office-agent-skills/main/install.sh | bash -s -- [options]
#
# Options:
#   --agent claude|codex|both     claude -> ~/.claude/skills (default); codex -> ~/.agents/skills
#   --skills-dir <path>           install into this folder instead (overrides --agent)
#   --firebase-config <file>      text file containing your firebaseConfig = {...}
#   --skip-deps                   do not run pip / npm
#   --check-only                  only report what is installed; change nothing

set -euo pipefail

REPO_URL="https://github.com/tedchen0814/office-agent-skills.git"
SKILLS=(html-slide-builder soil-html-deck soil-image-deck soil-teaching-deck draw file-toolkit)
PY_PACKAGES=(openai pillow python-pptx python-docx openpyxl pypdf reportlab pymupdf "qrcode[pil]" "markitdown[pdf,docx,pptx,xlsx]")
PPTXGENJS="pptxgenjs@3.12.0"

AGENT="claude"; SKILLS_DIR=""; FIREBASE_CFG=""; SKIP_DEPS=0; CHECK_ONLY=0
while [ $# -gt 0 ]; do
  case "$1" in
    --agent) AGENT="$2"; shift 2 ;;
    --skills-dir) SKILLS_DIR="$2"; shift 2 ;;
    --firebase-config) FIREBASE_CFG="$2"; shift 2 ;;
    --skip-deps) SKIP_DEPS=1; shift ;;
    --check-only) CHECK_ONLY=1; shift ;;
    -h|--help) sed -n '2,18p' "$0"; exit 0 ;;
    *) echo "unknown option: $1" >&2; exit 2 ;;
  esac
done

say()  { printf '\033[36m==> %s\033[0m\n' "$*"; }
ok()   { printf '\033[32m  [ok]   %s\033[0m\n' "$*"; }
warn() { printf '\033[33m  [warn] %s\033[0m\n' "$*"; }
bad()  { printf '\033[31m  [miss] %s\033[0m\n' "$*"; }
has()  { command -v "$1" >/dev/null 2>&1; }

find_python() {
  for c in python3 python; do
    if has "$c" && "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)' 2>/dev/null; then echo "$c"; return 0; fi
  done
  return 1
}

# ---------------------------------------------------------------- 1. prerequisites
say "Checking prerequisites"
HERE=""
if [ -n "${BASH_SOURCE[0]:-}" ] && [ -f "${BASH_SOURCE[0]}" ]; then HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"; fi
LOCAL_REPO=0; [ -n "$HERE" ] && [ -f "$HERE/skills/html-slide-builder/SKILL.md" ] && LOCAL_REPO=1

MISSING=()
if has git; then ok "git"; elif [ "$LOCAL_REPO" = 0 ]; then bad "git (needed to download the repo)"; MISSING+=(git); else warn "git not found (needed later by html-slide-builder's GitHub Pages deploy)"; fi
if PY="$(find_python)"; then ok "python: $PY"; else bad "python 3.9+"; MISSING+=(python); PY=""; fi
if has node; then ok "node $(node --version)"; else warn "node not found: soil-teaching-deck cannot build .pptx without it"; fi
if has gh;   then ok "gh (GitHub CLI)"; else warn "gh not found: html-slide-builder cannot publish decks to GitHub Pages"; fi

if [ ${#MISSING[@]} -gt 0 ] && [ "$CHECK_ONLY" = 0 ]; then
  echo; echo "Missing: ${MISSING[*]}" >&2
  echo "  macOS:  brew install ${MISSING[*]}" >&2
  echo "  Debian/Ubuntu:  sudo apt-get install -y ${MISSING[*]/python/python3 python3-pip}" >&2
  echo "Install them, then run this installer again." >&2
  exit 1
fi

# ---------------------------------------------------------------- 2. target folders
TARGETS=()
if [ -n "$SKILLS_DIR" ]; then TARGETS+=("$SKILLS_DIR")
else
  case "$AGENT" in
    claude) TARGETS+=("$HOME/.claude/skills") ;;
    codex)  TARGETS+=("$HOME/.agents/skills") ;;
    both)   TARGETS+=("$HOME/.claude/skills" "$HOME/.agents/skills") ;;
    *) echo "--agent must be claude, codex or both" >&2; exit 2 ;;
  esac
fi

key_ok() { [ -n "${OPENAI_API_KEY:-}" ] || [ -f "$HOME/.openai.env" ]; }

if [ "$CHECK_ONLY" = 1 ]; then
  say "Installed skills"
  for t in "${TARGETS[@]}"; do
    for s in "${SKILLS[@]}"; do if [ -f "$t/$s/SKILL.md" ]; then ok "$t/$s"; else bad "$t/$s"; fi; done
    fb="$t/html-slide-builder/references/firebase-config.md"
    if [ -f "$fb" ]; then if grep -q "YOUR_PROJECT_ID" "$fb"; then warn "Firebase not configured (live poll / word cloud slides will be skipped)"; else ok "Firebase configured"; fi; fi
  done
  if [ -n "$PY" ]; then if "$PY" -c "import openai, PIL, pptx, docx, openpyxl, pypdf, reportlab, pymupdf, qrcode, markitdown" 2>/dev/null; then ok "python packages"; else bad "python packages (${PY_PACKAGES[*]})"; fi; fi
  if has npm; then if npm ls -g pptxgenjs --depth=0 2>/dev/null | grep -q "pptxgenjs@"; then ok "pptxgenjs"; else bad "pptxgenjs (npm install -g $PPTXGENJS)"; fi; fi
  if key_ok; then ok "OpenAI key found (env or ~/.openai.env)"; else warn "no OPENAI_API_KEY yet: image generation will not work until you add one (see README)"; fi
  exit 0
fi

# ---------------------------------------------------------------- 3. get the repo
WORK="$(mktemp -d "${TMPDIR:-/tmp}/office-agent-skills.XXXXXX")"
trap 'rm -rf "$WORK"' EXIT
BACKUP_ROOT="$HOME/.claude/skill-backups/$(date +%Y%m%d-%H%M%S)"

say "Getting office-agent-skills"
if [ "$LOCAL_REPO" = 1 ]; then REPO="$HERE"; ok "using local copy: $REPO"
else
  REPO="$WORK/office-agent-skills"
  git -c core.longpaths=true clone --quiet --depth 1 "$REPO_URL" "$REPO"
  ok "cloned $REPO_URL"
fi

# ---------------------------------------------------------------- 4. copy skills + resolve paths
say "Installing skills"
for t in "${TARGETS[@]}"; do
  mkdir -p "$t"
  for s in "${SKILLS[@]}"; do
    if [ -d "$t/$s" ]; then
      mkdir -p "$BACKUP_ROOT"
      mv "$t/$s" "$BACKUP_ROOT/$(basename "$(dirname "$t")")-$s"
    fi
    cp -R "$REPO/skills/$s" "$t/$s"
    ok "$s -> $t/$s"
  done
  POST=(--skills-dir "$t"); [ -n "$FIREBASE_CFG" ] && POST+=(--firebase-config "$FIREBASE_CFG")
  "$PY" "$REPO/installer/postinstall.py" "${POST[@]}"
done
[ -d "$BACKUP_ROOT" ] && warn "previous versions were moved to $BACKUP_ROOT"

# ---------------------------------------------------------------- 5. dependencies
if [ "$SKIP_DEPS" = 0 ]; then
  say "Installing Python packages: ${PY_PACKAGES[*]}"
  "$PY" -m pip install --quiet --disable-pip-version-check "${PY_PACKAGES[@]}" \
    || "$PY" -m pip install --quiet --disable-pip-version-check --break-system-packages "${PY_PACKAGES[@]}"
  ok "python packages"
  if has npm; then
    say "Installing $PPTXGENJS (global)"
    if npm install -g "$PPTXGENJS"; then ok "$PPTXGENJS"; else warn "npm install failed (permissions?). Fix npm's global prefix instead of using sudo, then run: npm install -g $PPTXGENJS"; fi
  else warn "npm not found; skipped $PPTXGENJS"; fi
fi

# ---------------------------------------------------------------- 6. verify
say "Environment check"
for t in "${TARGETS[@]}"; do for s in "${SKILLS[@]}"; do if [ -f "$t/$s/SKILL.md" ]; then ok "$s"; else bad "$s (in $t)"; fi; done; done
if key_ok; then ok "OpenAI key found"; else
  warn "No OpenAI key yet. Image generation needs YOUR OWN key (billed to your account):"
  echo "         create ~/.openai.env containing one line:  OPENAI_API_KEY=sk-..."
  echo "         and complete Individual verification at platform.openai.com/settings/organization/general"
fi
echo
printf '\033[32mDone. Restart your agent (Claude Code / Codex) so it loads the new skills.\033[0m\n'
