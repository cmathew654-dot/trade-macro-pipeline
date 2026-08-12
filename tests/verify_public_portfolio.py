#!/usr/bin/env python3
"""Check the tracked public tree without walking the worktree."""

from pathlib import Path
import re
import subprocess

ROOT = Path(__file__).resolve().parents[1]
ARTIFACT = re.compile(r"^(?:\.agents|\.claude|\.codex|\.planning|\.ship|\.superpowers|docs/(?:codex|superpowers))(?:/|$)")

tracked = subprocess.run(
    ["git", "ls-files"], cwd=ROOT, check=True, capture_output=True, text=True
).stdout.splitlines()
failures = [path for path in tracked if ARTIFACT.match(path)]

for path in tracked:
    if path.endswith(".md") and "summit financial group" in (ROOT / path).read_text(encoding="utf-8", errors="replace").lower():
        failures.append(f"{path} names an employer")

if failures:
    raise AssertionError("\n".join(failures))
print("Public portfolio checks passed.")
