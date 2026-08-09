#!/usr/bin/env python3
"""Static checks for public fixture and README boundaries."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEXT_SUFFIXES = {".bas", ".cls", ".frm", ".md", ".py", ".ps1", ".xml", ".csv", ".txt"}
FORBIDDEN_EMPLOYERS = ("summit financial group",)


def main() -> None:
    failures: list[str] = []
    for path in ROOT.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in TEXT_SUFFIXES:
            continue
        if path.resolve() == Path(__file__).resolve():
            continue
        if ".git" in path.parts or "__pycache__" in path.parts:
            continue
        text = path.read_text(encoding="utf-8", errors="replace").lower()
        for phrase in FORBIDDEN_EMPLOYERS:
            if phrase in text:
                failures.append(f"{path.relative_to(ROOT)} contains {phrase!r}")

    readme = (ROOT / "README.md").read_text(encoding="utf-8").lower()
    for phrase in ("human-gated", "10,800_lines", "8_scripts", "practicing financial advisor"):
        if phrase in readme:
            failures.append(f"README.md contains stale portfolio phrase {phrase!r}")

    if failures:
        raise AssertionError("\n".join(failures))
    print("Public portfolio checks passed.")


if __name__ == "__main__":
    main()
