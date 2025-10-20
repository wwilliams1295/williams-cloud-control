# core/execute.py — safely run the planned steps
from typing import Any
from core.capabilities import REGISTRY


def run_plan(plan: list[dict[str, Any]]) -> dict[str, Any]:
    last = {}
    for step in plan:
        tool = step.get("tool")
        args = step.get("args", {})
        if tool not in [s.name for s in REGISTRY.list_specs()]:
            return {"ok": False, "error": f"Unknown tool: {tool}"}
        spec = REGISTRY.get(tool)
        out = spec.entrypoint(**args)
        if not out or not out.get("ok", False):
            return {"ok": False, "error": f"{tool} failed", "detail": out}
        last = out
    return {"ok": True, "result": last}
