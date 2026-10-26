#!/usr/bin/env python3
"""Post-install step shared by install.ps1 and install.sh.

1. Replace the {{SKILLS_DIR}} placeholder in the installed skills with the real
   skills folder (forward slashes), so draw.py / pack_pptx.py paths resolve.
2. Optionally fill the YOUR_* placeholders in html-slide-builder's
   references/firebase-config.md from a pasted firebaseConfig snippet.

Usage:
  python postinstall.py --skills-dir ~/.claude/skills [--firebase-config cfg.js]
"""
import argparse
import re
import sys
from pathlib import Path

SKILLS = ["html-slide-builder", "soil-html-deck", "soil-image-deck", "soil-teaching-deck", "draw", "file-toolkit"]
TEXT_EXT = {".md", ".py", ".json", ".txt", ".html", ".js"}
FIREBASE_KEYS = {
    "apiKey": "YOUR_FIREBASE_API_KEY",
    "messagingSenderId": "YOUR_MESSAGING_SENDER_ID",
    "appId": "YOUR_APP_ID",
}


def read(p: Path) -> str:
    return p.read_text(encoding="utf-8")


def write(p: Path, s: str) -> None:
    p.write_text(s, encoding="utf-8", newline="\n")


def substitute_paths(skills_dir: Path) -> int:
    value = skills_dir.resolve().as_posix()
    n = 0
    for name in SKILLS:
        root = skills_dir / name
        if not root.is_dir():
            continue
        for f in root.rglob("*"):
            if f.is_file() and f.suffix.lower() in TEXT_EXT:
                try:
                    s = read(f)
                except UnicodeDecodeError:
                    continue
                if "{{SKILLS_DIR}}" in s:
                    write(f, s.replace("{{SKILLS_DIR}}", value))
                    n += 1
    return n


def parse_firebase(text: str) -> dict:
    out = {}
    for key in ("apiKey", "authDomain", "projectId", "storageBucket", "messagingSenderId", "appId"):
        m = re.search(r"""\b%s\s*:\s*["']([^"']+)["']""" % key, text)
        if m:
            out[key] = m.group(1)
    missing = [k for k in ("apiKey", "projectId", "messagingSenderId", "appId") if k not in out]
    if missing:
        raise ValueError("firebaseConfig is missing: " + ", ".join(missing))
    return out


def fill_firebase(skills_dir: Path, cfg_file: Path) -> bool:
    target = skills_dir / "html-slide-builder" / "references" / "firebase-config.md"
    if not target.is_file():
        print("  [warn] html-slide-builder is not installed; skipping Firebase config")
        return False
    cfg = parse_firebase(read(cfg_file))
    pid = cfg["projectId"]
    s = read(target)
    s = s.replace("YOUR_PROJECT_ID.firebaseapp.com", cfg.get("authDomain", pid + ".firebaseapp.com"))
    s = s.replace("YOUR_PROJECT_ID.firebasestorage.app", cfg.get("storageBucket", pid + ".firebasestorage.app"))
    s = s.replace("YOUR_PROJECT_ID", pid)
    for key, placeholder in FIREBASE_KEYS.items():
        s = s.replace(placeholder, cfg[key])
    write(target, s)
    return True


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--skills-dir", required=True)
    ap.add_argument("--firebase-config")
    a = ap.parse_args()
    skills_dir = Path(a.skills_dir).expanduser()

    n = substitute_paths(skills_dir)
    print("  [ok]   paths resolved in %d file(s) -> %s" % (n, skills_dir.resolve().as_posix()))

    fb = skills_dir / "html-slide-builder" / "references" / "firebase-config.md"
    if a.firebase_config:
        try:
            if fill_firebase(skills_dir, Path(a.firebase_config).expanduser()):
                print("  [ok]   Firebase config written to html-slide-builder")
        except (ValueError, OSError) as e:
            print("  [warn] Firebase config not applied: %s" % e)
    elif fb.is_file() and "YOUR_PROJECT_ID" in read(fb):
        print("  [note] Firebase not configured: html-slide-builder will skip live word cloud / poll slides")
        print("         until you give it your own firebaseConfig (see README).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
