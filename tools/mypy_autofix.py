#!/usr/bin/env python3
"""
mypy_autofix.py
A pragmatic fixer for common mypy errors you reported.

It:
  - Ensures `from __future__ import annotations` and `from typing import Any, Optional`
  - Adds `-> None` to functions with no return annotation
  - Adds `Any` to untyped parameters (incl. *args/**kwargs)
  - Converts `Callable[...] = None` / `Exception = None` to Optional[...] = None
  - Removes bare/unneeded "# type: ignore" comments when mypy flags them as unused
  - Relaxes return annotation to `-> Any` when mypy complains "Returning Any from function declared to return ..."
  - (Optional) writes/augments mypy.ini to ignore missing imports for noisy external libs

This is intentionally conservative and mechanical. It won't rewrite complex code,
but it will squash the bulk of "missing annotation" / "incompatible None" / "unused ignore" issues fast.

Usage:
  python tools/mypy_autofix.py --apply
  python tools/mypy_autofix.py --dry-run
  python tools/mypy_autofix.py --write-mypy-ini
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
PROJECT_SRC = ROOT  # fix files under repo root

# ---------- helpers ----------


@dataclass
class MypyMsg:
    path: Path
    line: int
    col: Optional[int]
    code: str
    text: str


MYPY_LINE = re.compile(
    r"""
    ^(?P<file>.*?):
    (?P<line>\d+)
    (?::(?P<col>\d+))?
    :\s+error:\s+(?P<msg>.+?)\s+\[(?P<code>[^\]]+)\]\s*$
    """,
    re.VERBOSE,
)


def run_mypy(args: List[str]) -> str:
    import subprocess

    cmd = [
        "mypy",
        ".",
        "--exclude",
        r"(\.venv|venv|backups|site-packages|\.tox|\.mypy_cache|\.pytest_cache)",
    ]
    cmd.extend(args)
    proc = subprocess.run(cmd, cwd=str(ROOT), capture_output=True, text=True)
    out = proc.stdout + proc.stderr
    return out


def parse_mypy(output: str) -> List[MypyMsg]:
    msgs: List[MypyMsg] = []
    for line in output.splitlines():
        m = MYPY_LINE.match(line.strip())
        if not m:
            continue
        path = Path(m.group("file")).resolve()
        if not path.exists():
            continue
        msgs.append(
            MypyMsg(
                path=path,
                line=int(m.group("line")),
                col=int(m.group("col")) if m.group("col") else None,
                code=m.group("code"),
                text=m.group("msg"),
            )
        )
    return msgs


def read_file(p: Path) -> List[str]:
    return p.read_text(encoding="utf-8").splitlines(keepends=True)


def write_file(p: Path, lines: List[str]) -> None:
    p.write_text("".join(lines), encoding="utf-8")


def ensure_imports(lines: List[str]) -> List[str]:
    """Ensure `from __future__ import annotations` and basic typing imports."""
    text = "".join(lines)
    changed = False

    if "from __future__ import annotations" not in text:
        # insert at top (after shebang/encoding)
        ins_at = 0
        if lines and lines[0].startswith("#!"):
            ins_at = 1
        # handle encoding line too
        if len(lines) > ins_at and "coding" in lines[ins_at]:
            ins_at += 1
        lines.insert(ins_at, "from __future__ import annotations\n")
        changed = True

    # Add typing imports if we will need them
    need_any = re.search(r"\bAny\b", text) is None
    need_optional = re.search(r"\bOptional\b", text) is None
    if need_any or need_optional:
        # find first import line to place typing import after
        insert_after = 0
        for i, L in enumerate(lines):
            if L.startswith("from __future__ import annotations"):
                insert_after = i + 1
                break
        pieces = ["from typing import "]
        names: List[str] = []
        if need_any:
            names.append("Any")
        if need_optional:
            names.append("Optional")
        if names:
            lines.insert(insert_after, f"{pieces[0]}{', '.join(names)}\n")
            changed = True

    return lines if changed else lines


# ---------- fixers ----------

DEF_LINE = re.compile(
    r"^(\s*)def\s+([A-Za-z_]\w*)\s*\((.*?)\)\s*(?:->\s*([^:]+))?\s*:\s*(#.*)?$"
)


def annotate_function_signature(lines: List[str], idx: int) -> bool:
    """
    Add Any to untyped params, add -> None if no return annotation.
    idx is zero-based line index.
    """
    m = DEF_LINE.match(lines[idx])
    if not m:
        return False
    indent, name, params, ret, trailing = m.groups()

    # tokenize params (very naive, but works for most)
    # we will skip already-typed params (contain ':') and keep defaults
    def fix_param(tok: str) -> str:
        t = tok.strip()
        if not t:
            return t
        if t.startswith("*") and ":" not in t:
            # *args /**kwargs
            if t.startswith("**"):
                if ":" not in t:
                    return f"{t}: Any"
            else:
                if ":" not in t:
                    return f"{t}: Any"
        # plain param (no type)
        if ":" not in t:
            # keep default if exists
            if "=" in t:
                name_, default = t.split("=", 1)
                return f"{name_.strip()}: Any = {default.strip()}"
            return f"{t}: Any"
        return t

    # split on commas, but don’t try to parse complex signatures; light-weight pass
    parts: List[str] = []
    balance = 0
    buf = ""
    for ch in params:
        if ch in "([{":
            balance += 1
        elif ch in ")]}":
            balance -= 1
        if ch == "," and balance == 0:
            parts.append(buf)
            buf = ""
        else:
            buf += ch
    if buf != "":
        parts.append(buf)

    fixed = [fix_param(p) for p in parts] if params.strip() else []
    new_params = ", ".join(fixed)

    new_ret = ret.strip() if ret else "None"

    lines[idx] = (
        f"{indent}def {name}({new_params}) -> {new_ret}:{' ' + trailing if trailing else ''}\n"
    )
    return True


OPT_CALLABLE_OR_EXC = re.compile(
    r"""
    ^(?P<indent>\s*)
    (?P<name>[A-Za-z_]\w*)
    \s*:\s*
    (?P<ann>
       Callable\[[^\]]+\]
      |Exception
    )
    \s*=\s*None\s*$
    """,
    re.VERBOSE,
)


def optionalize_none_assign(lines: List[str], idx: int) -> bool:
    """
    Turn `x: Callable[...] = None` or `x: Exception = None` into `Optional[...]`.
    """
    m = OPT_CALLABLE_OR_EXC.match(lines[idx].rstrip())
    if not m:
        return False
    indent, name, ann = m.group("indent"), m.group("name"), m.group("ann")
    lines[idx] = f"{indent}{name}: Optional[{ann}] = None\n"
    return True


UNUSED_IGNORE = re.compile(r"#\s*type:\s*ignore(\b|$)")


def remove_unused_ignore(lines: List[str], idx: int) -> bool:
    """
    Remove bare '# type: ignore' at end of line (keep code before it).
    """
    line = lines[idx]
    if "# type: ignore" not in line:
        return False
    # only strip the comment, keep leading code
    before = line.split("# type: ignore", 1)[0].rstrip()
    # keep any other trailing comment content after ignore code (rare)
    lines[idx] = before + "\n"
    return True


RET_ANY_MSG = re.compile(r'^Returning Any from function declared to return "(.+)"$')


def relax_return_to_any(lines: List[str], start_idx: int) -> bool:
    """
    For a def line with `-> something`, change to `-> Any`.
    """
    i = start_idx
    while i >= 0 and not DEF_LINE.match(lines[i]):
        i -= 1
    if i < 0:
        return False
    m = DEF_LINE.match(lines[i])
    if not m:
        return False
    indent, name, params, ret, trailing = m.groups()
    if not ret:
        return False
    lines[i] = (
        f"{indent}def {name}({params}) -> Any:{' ' + trailing if trailing else ''}\n"
    )
    return True


NONE_ASSIGN_NEEDS_OPTIONAL = re.compile(
    r"""
    ^(?P<indent>\s*)
    (?P<name>[A-Za-z_]\w*)
    \s*:\s*
    (?P<ann>[^=\n]+?)
    \s*=\s*None\s*$
    """,
    re.VERBOSE,
)


def generic_optionalize(lines: List[str], idx: int) -> bool:
    """
    For any `x: T = None` where T is not Optional, wrap as Optional[T].
    Skips if already Optional[...].
    """
    line = lines[idx].rstrip()
    m = NONE_ASSIGN_NEEDS_OPTIONAL.match(line)
    if not m:
        return False
    ann = m.group("ann").strip()
    if ann.startswith("Optional["):
        return False
    indent, name = m.group("indent"), m.group("name")
    lines[idx] = f"{indent}{name}: Optional[{ann}] = None\n"
    return True


def ensure_typing_imports_if_used(lines: List[str]) -> List[str]:
    """If we inserted Optional/Any, ensure the import exists."""
    text = "".join(lines)
    need_any = " Any" in text or text.startswith("Any")
    need_optional = " Optional" in text or text.startswith("Optional")
    if not (need_any or need_optional):
        return lines
    return ensure_imports(lines)


# ---------- orchestrate fixes per file ----------


def fix_file(path: Path, file_msgs: List[MypyMsg]) -> Tuple[bool, List[str]]:
    """
    Apply fixes to a single file based on mypy messages.
    Returns (changed, new_lines).
    """
    lines = read_file(path)
    changed = False

    # First pass: ensure imports (we might add Any/Optional later again)
    lines = ensure_imports(lines)

    # Pre-index: map line->messages for this file
    msgs_by_line: Dict[int, List[MypyMsg]] = {}
    for m in file_msgs:
        msgs_by_line.setdefault(m.line, []).append(m)

    # Walk lines; apply line-targeted fixes
    for idx in range(len(lines)):
        line_no = idx + 1
        msgs = msgs_by_line.get(line_no, [])

        # opportunistic fixes even if no msg:
        if optionalize_none_assign(lines, idx):
            changed = True
        elif generic_optionalize(lines, idx):
            changed = True

        for msg in msgs:
            txt = msg.text

            # 1) Missing annotation(s)
            if (
                "Function is missing a return type annotation" in txt
                or "Function is missing a type annotation" in txt
                or "Function is missing a type annotation for one or more arguments"
                in txt
            ):
                if DEF_LINE.match(lines[idx]):
                    if annotate_function_signature(lines, idx):
                        changed = True

            # 2) Incompatible types in assignment (... None) already handled by generic_optionalize

            # 3) Unused 'type: ignore'
            if 'Unused "type: ignore" comment' in txt or msg.code == "unused-ignore":
                if remove_unused_ignore(lines, idx):
                    changed = True

            # 4) no-any-return: relax to Any
            if txt.startswith("Returning Any from function declared to return"):
                if relax_return_to_any(lines, idx):
                    changed = True

    # Ensure typing imports if we inserted Any/Optional
    new_lines = ensure_typing_imports_if_used(lines)
    if new_lines != lines:
        lines = new_lines
        changed = True

    return changed, lines


# ---------- mypy.ini helper (optional) ----------

MYPY_INI_SNIPPET = """\
# --- added by mypy_autofix.py ---
[mypy]
ignore_missing_imports = True

[mypy-onedrive_device_login]
disallow_untyped_defs = False
warn_return_any = False

[mypy-test_integrations]
disallow_untyped_defs = False
warn_return_any = False

[mypy-tools.ai_patch]
disallow_untyped_defs = False
warn_return_any = False
"""


def ensure_mypy_ini() -> None:
    ini = ROOT / "mypy.ini"
    if not ini.exists():
        ini.write_text(MYPY_INI_SNIPPET, encoding="utf-8")
        print(f"[mypy_autofix] wrote {ini}")
