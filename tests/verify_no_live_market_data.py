"""Regression boundary: the current CDS pipeline contains no live-market-data integration."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCAN_ROOTS = ("vba", "ribbon", "deploy", "tests")
TEXT_SUFFIXES = {".bas", ".cls", ".frm", ".xml", ".ps1", ".py", ".md"}
BANNED = tuple("".join(parts).lower() for parts in (
    ("CDS_", "PriceGuard"), ("Refresh", "LivePrices"),
    ("refresh", "_prices"), ("ConvertTo", "LinkedDataType"),
    ("FIELD", "VALUE"), ("CDS Live", " Prices"),
    ("DriftAlert", "Pct"), ("Excel ", "Stocks"),
    ("linked data", " type"), ("live ", "price"), ("live ", "quote"),
))


def candidate_files():
    files = [ROOT / "README.md"]
    for directory in SCAN_ROOTS:
        files.extend(path for path in (ROOT / directory).rglob("*")
                     if path.is_file() and path.suffix.lower() in TEXT_SUFFIXES)
    return sorted(files)


def main():
    violations = []
    for path in candidate_files():
        text = path.read_text(encoding="utf-8", errors="replace").lower()
        for token in BANNED:
            if token in text:
                violations.append(f"{path.relative_to(ROOT)}: {token}")
    updater = (ROOT / "deploy" / "Update-PersonalXlsb.ps1").read_text(
        encoding="utf-8", errors="replace")
    for required in ("$obsoleteModuleNames", "$proj.VBComponents.Remove($obsolete)"):
        if required not in updater:
            violations.append(f"deploy/Update-PersonalXlsb.ps1: missing {required}")
    if violations:
        print("FAIL: live-market-data boundary violations")
        for violation in violations:
            print(f"  {violation}")
        return 1
    print("PASS: no live-market-data integration remains in the current pipeline")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
