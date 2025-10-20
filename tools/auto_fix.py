# tools/auto_fix.py
import re
import sys
from pathlib import Path
from typing import Iterable

REPO = Path(__file__).resolve().parents[1]
EXCLUDES = {
    ".venv",
    "venv",
    "backups",
    "site-packages",
    ".tox",
    ".mypy_cache",
    ".pytest_cache",
}


def should_format_file(p: Path) -> bool:
    if p.suffix != ".py":
        return False
    if any(ex in p.parts for ex in EXCLUDES):
        return False
    return True


def fix_semicolons_and_one_liners(text: str) -> str:
    """Manual repair for E701/E702/E703 only."""
    lines = text.splitlines()

    def split_semicolons(line: str) -> list[str]:
        if ";" not in line:
            return [line]
        # Simple heuristic: split only when not inside string
        if '"' in line or "'" in line:
            return [line]
        parts = [seg.rstrip() for seg in line.split(";") if seg.strip()]
        return parts if len(parts) > 1 else [line]

    def expand_one_line_blocks(line: str) -> list[str]:
        patterns = [
            r"^\s*(if\s+.+?:)\s+(.*\S.*)$",
            r"^\s*(elif\s+.+?:)\s+(.*\S.*)$",
            r"^\s*(else\s*:)\s+(.*\S.*)$",
            r"^\s*(for\s+.+?:)\s+(.*\S.*)$",
            r"^\s*(while\s+.+?:)\s+(.*\S.*)$",
            r"^\s*(with\s+.+?:)\s+(.*\S.*)$",
            r"^\s*(try\s*:)\s+(.*\S.*)$",
            r"^\s*(except\s+.*?:)\s+(.*\S.*)$",
            r"^\s*(finally\s*:)\s+(.*\S.*)$",
        ]
        for pat in patterns:
            m = re.match(pat, line)
            if m:
                head, tail = m.groups()
                indent = " " * (len(line) - len(line.lstrip()))
                return [f"{indent}{head}", f"{indent}    {tail}"]
        return [line]

    # drop trailing semicolons
    lines = [re.sub(r";\s*$", "", l) for l in lines]

    new_lines = []
    for ln in lines:
        for part in split_semicolons(ln):
            new_lines.extend(expand_one_line_blocks(part))

    return "\n".join(new_lines)


def rewrite_files(paths: Iterable[Path]) -> None:
    for p in paths:
        if not should_format_file(p):
            continue
        try:
            src = p.read_text(encoding="utf-8")
            fixed = fix_semicolons_and_one_liners(src)
            if fixed != src:
                p.write_text(fixed, encoding="utf-8")
                print(f"[auto-fix] Cleaned {p}")
        except Exception as e:
            print(f"[auto-fix] Skip {p}: {e}")


def main() -> int:
    # Just our own repair logic — no Ruff or Bandit calls
    py_files = [p for p in REPO.rglob("*.py") if should_format_file(p)]
    rewrite_files(py_files)
    print("[auto-fix] Completed manual cleanup pass.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
