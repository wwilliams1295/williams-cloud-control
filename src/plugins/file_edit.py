from typing import Any
from pathlib import Path

name = "file_edit"
description = "Find/replace text in a file (writes .bak)"
permissions: list[str] = ["fs:read:*", "fs:write:*"]


def run(path: str, find: str, replace: str, backup: bool = True) -> dict[str, Any]:
    p = Path(path)
    data = p.read_text(encoding="utf-8")
    if backup:
        (p.parent / (p.name + ".bak")).write_text(data, encoding="utf-8")
    new = data.replace(find, replace)
    p.write_text(new, encoding="utf-8")
    return {"ok": True, "replacements": data.count(find)}
