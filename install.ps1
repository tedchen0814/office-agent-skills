# work-slide-skills one-step installer (Windows PowerShell 5.1+ / PowerShell 7)
#
# Installs these skills into the agent's skills folder:
#   html-slide-builder   Reveal.js HTML deck + AI backgrounds/icons + Firebase poll/word cloud + GitHub Pages
#   soil-html-deck       single-file HTML deck (Chart.js, tables, video)
#   soil-image-deck      every slide is one AI-generated image, packed into .pptx
#   soil-teaching-deck   editable .pptx with AI illustrations (PptxGenJS)
#   draw                 gpt-image-2 image generation helper used by the four above
#   file-toolkit         how-to for Word / Excel / PowerPoint / PDF batch work, QR codes, file-to-Markdown
# and the Python / npm packages they need. It never installs or stores an OpenAI key.
#
# Usage (from a clone):   powershell -ExecutionPolicy Bypass -File .\install.ps1 [options]
# Usage (no clone):       & ([scriptblock]::Create((irm https://raw.githubusercontent.com/tedchen0814/work-slide-skills/main/install.ps1))) [options]
#
# Options:
#   -Agent claude|codex|both    claude -> ~\.claude\skills (default); codex -> ~\.agents\skills
#   -SkillsDir <path>           install into this folder instead (overrides -Agent)
#   -FirebaseConfigFile <file>  text file containing your firebaseConfig = {...}; fills html-slide-builder
#   -InstallPrereqs             install missing Git / Python / Node.js with winget
#   -SkipDeps                   do not run pip / npm
#   -CheckOnly                  only report what is installed; change nothing

[CmdletBinding()]
param(
  [ValidateSet("claude", "codex", "both")] [string]$Agent = "claude",
  [string]$SkillsDir = "",
  [string]$FirebaseConfigFile = "",
  [switch]$InstallPrereqs,
  [switch]$SkipDeps,
  [switch]$CheckOnly
)

$ErrorActionPreference = "Stop"

$RepoUrl    = "https://github.com/tedchen0814/work-slide-skills.git"
$Skills     = @("html-slide-builder", "soil-html-deck", "soil-image-deck", "soil-teaching-deck", "draw", "file-toolkit")
$PyPackages = @("openai", "pillow", "python-pptx", "python-docx", "openpyxl", "pypdf", "reportlab", "pymupdf", "qrcode[pil]", "markitdown[pdf,docx,pptx,xlsx]")
$PptxGenJs  = "pptxgenjs@3.12.0"

function Say($msg)  { Write-Host "==> $msg" -ForegroundColor Cyan }
function Ok($msg)   { Write-Host "  [ok]   $msg" -ForegroundColor Green }
function Warn($msg) { Write-Host "  [warn] $msg" -ForegroundColor Yellow }
function Bad($msg)  { Write-Host "  [miss] $msg" -ForegroundColor Red }
function Has($cmd)  { return [bool](Get-Command $cmd -ErrorAction SilentlyContinue) }
function Refresh-Path {
  $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
}

function Find-Python {
  # The Microsoft Store "python" stub exits with code 49, so test each candidate for real.
  foreach ($c in @(@("py", "-3"), @("python"), @("python3"))) {
    if (-not (Has $c[0])) { continue }
    try {
      $a = @(); if ($c.Count -gt 1) { $a += $c[1..($c.Count - 1)] }
      $v = & $c[0] @a -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
      if ($LASTEXITCODE -eq 0 -and $v -match '^3\.(\d+)$' -and [int]$Matches[1] -ge 9) { return ,$c }
    } catch { }
  }
  return $null
}

# ---------------------------------------------------------------- 1. prerequisites
Say "Checking prerequisites"
$missing = @()
$needGit = -not ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot "skills\html-slide-builder\SKILL.md")))
if (Has "git") { Ok "git" } elseif ($needGit) { Bad "git (needed to download the repo)"; $missing += "git" } else { Warn "git not found (needed later by html-slide-builder's GitHub Pages deploy)" }
$py = Find-Python
if ($py) { Ok "python: $($py -join ' ')" } else { Bad "python 3.9+"; $missing += "python" }
if (Has "node") { Ok "node $(& node --version)" } else { Warn "node not found: soil-teaching-deck cannot build .pptx without it" }
if (Has "gh")   { Ok "gh (GitHub CLI)" } else { Warn "gh not found: html-slide-builder can build decks but cannot publish them to GitHub Pages" }

$winget = @{ git = "Git.Git"; python = "Python.Python.3.12" }
if ($missing.Count -gt 0 -and -not $CheckOnly) {
  if ($InstallPrereqs -and (Has "winget")) {
    foreach ($m in $missing) { Say "Installing $m with winget"; & winget install --id $winget[$m] -e --accept-source-agreements --accept-package-agreements }
    Refresh-Path
    $py = Find-Python
    if (-not $py) { Warn "Open a NEW terminal and run this installer again."; exit 1 }
  } else {
    Write-Host ""
    Write-Host "Missing: $($missing -join ', ')" -ForegroundColor Red
    foreach ($m in $missing) { Write-Host "  winget install --id $($winget[$m]) -e" }
    Write-Host "Or rerun with -InstallPrereqs. Then run this installer again."
    exit 1
  }
}

# ---------------------------------------------------------------- 2. target folders
$targets = @()
if ($SkillsDir) { $targets += $SkillsDir }
else {
  if ($Agent -in @("claude", "both")) { $targets += (Join-Path $HOME ".claude\skills") }
  if ($Agent -in @("codex", "both"))  { $targets += (Join-Path $HOME ".agents\skills") }
}

if ($CheckOnly) {
  Say "Installed skills"
  foreach ($t in $targets) {
    foreach ($s in $Skills) { if (Test-Path (Join-Path $t "$s\SKILL.md")) { Ok "$t\$s" } else { Bad "$t\$s" } }
    $fb = Join-Path $t "html-slide-builder\references\firebase-config.md"
    if (Test-Path $fb) {
      if ((Get-Content $fb -Raw -Encoding UTF8) -match "YOUR_PROJECT_ID") { Warn "Firebase not configured (live poll / word cloud slides will be skipped)" } else { Ok "Firebase configured" }
    }
  }
  if ($py) {
    $a = @(); if ($py.Count -gt 1) { $a += $py[1..($py.Count - 1)] }
    & $py[0] @a -c "import openai, PIL, pptx, docx, openpyxl, pypdf, reportlab, pymupdf, qrcode, markitdown" 2>$null
    if ($LASTEXITCODE -eq 0) { Ok "python packages" } else { Bad "python packages ($($PyPackages -join ', '))" }
  }
  if (Has "npm") { $g = (& npm ls -g pptxgenjs --depth=0 2>$null) -join ' '; if ($g -match "pptxgenjs@") { Ok "pptxgenjs" } else { Bad "pptxgenjs (npm install -g $PptxGenJs)" } }
  $keyOk = [bool]$env:OPENAI_API_KEY -or (Test-Path (Join-Path $HOME ".openai.env"))
  if ($keyOk) { Ok "OpenAI key found (env or ~/.openai.env)" } else { Warn "no OPENAI_API_KEY yet: image generation will not work until you add one (see README)" }
  exit 0
}

# ---------------------------------------------------------------- 3. get the repo
$work = Join-Path ([IO.Path]::GetTempPath()) ("work-slide-skills-" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Force -Path $work | Out-Null
# Backups go OUTSIDE the skills folder, or the agent would load two copies of a skill.
$backupRoot = Join-Path $HOME (".claude\skill-backups\" + (Get-Date -Format "yyyyMMdd-HHmmss"))

try {
  Say "Getting work-slide-skills"
  if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot "skills\html-slide-builder\SKILL.md"))) {
    $repo = $PSScriptRoot
    Ok "using local copy: $repo"
  } else {
    $repo = Join-Path $work "work-slide-skills"
    & git -c core.longpaths=true clone --quiet --depth 1 $RepoUrl $repo
    if ($LASTEXITCODE -ne 0) { throw "git clone $RepoUrl failed" }
    Ok "cloned $RepoUrl"
  }

  # ---------------------------------------------------------------- 4. copy skills + resolve paths
  Say "Installing skills"
  foreach ($t in $targets) {
    New-Item -ItemType Directory -Force -Path $t | Out-Null
    foreach ($s in $Skills) {
      $dst = Join-Path $t $s
      if (Test-Path $dst) {
        New-Item -ItemType Directory -Force -Path $backupRoot | Out-Null
        Move-Item -Path $dst -Destination (Join-Path $backupRoot ((Split-Path (Split-Path $t -Parent) -Leaf) + "-" + $s))
      }
      Copy-Item -Recurse -Path (Join-Path $repo "skills\$s") -Destination $dst
      Ok "$s -> $dst"
    }
    $pa = @(); if ($py.Count -gt 1) { $pa += $py[1..($py.Count - 1)] }
    $post = @("--skills-dir", $t)
    if ($FirebaseConfigFile) { $post += @("--firebase-config", $FirebaseConfigFile) }
    & $py[0] @pa (Join-Path $repo "installer\postinstall.py") @post
    if ($LASTEXITCODE -ne 0) { throw "postinstall.py failed" }
  }
  if (Test-Path $backupRoot) { Warn "previous versions were moved to $backupRoot" }

  # ---------------------------------------------------------------- 5. dependencies
  if (-not $SkipDeps) {
    Say "Installing Python packages: $($PyPackages -join ', ')"
    $pa = @(); if ($py.Count -gt 1) { $pa += $py[1..($py.Count - 1)] }
    & $py[0] @pa -m pip install --quiet --disable-pip-version-check @PyPackages
    if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
    Ok "python packages"
    if (Has "npm") {
      Say "Installing $PptxGenJs (global)"
      & npm install -g $PptxGenJs
      if ($LASTEXITCODE -ne 0) { Warn "npm install failed; soil-teaching-deck will not be able to build .pptx" } else { Ok $PptxGenJs }
    } else { Warn "npm not found; skipped $PptxGenJs" }
  }

  # ---------------------------------------------------------------- 6. verify
  Say "Environment check"
  foreach ($t in $targets) { foreach ($s in $Skills) { if (Test-Path (Join-Path $t "$s\SKILL.md")) { Ok $s } else { Bad "$s (in $t)" } } }
  $keyOk = [bool]$env:OPENAI_API_KEY -or (Test-Path (Join-Path $HOME ".openai.env"))
  if ($keyOk) { Ok "OpenAI key found" } else {
    Warn "No OpenAI key yet. Image generation needs YOUR OWN key (billed to your account):"
    Write-Host "         create ~/.openai.env containing one line:  OPENAI_API_KEY=sk-..."
    Write-Host "         and complete Individual verification at platform.openai.com/settings/organization/general"
  }
  Write-Host ""
  Write-Host "Done. Restart your agent (Claude Code / Codex) so it loads the new skills, then try:" -ForegroundColor Green
  Write-Host "  describe the deck you want, e.g. 'make an interactive HTML deck from this outline' (html-slide-builder)"
} finally {
  Remove-Item -Recurse -Force -Path $work -ErrorAction SilentlyContinue
}
