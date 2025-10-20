import importlib
import pkgutil
import typing as t
from . import Plugin

_REGISTRY: dict[str, Plugin] = {}


def load_all(package_name: str = "plugins") -> dict[str, Plugin]:
    global _REGISTRY
    _REGISTRY.clear()
    pkg = importlib.import_module(package_name)
    for m in pkgutil.iter_modules(pkg.__path__):
        if m.name in {"__init__", "loader", "init"}:
            continue
        modname = f"{package_name}.{m.name}"
        try:
            mod = importlib.import_module(modname)
        except Exception as e:
            print(f"[plugins] ⚠️ failed to import {modname}: {e}")
            continue
        missing = [k for k in ("name", "description", "run") if not hasattr(mod, k)]
        if missing:
            print(f"[plugins] ⏭️ skipping {modname}: missing {missing}")
            continue
        _REGISTRY[getattr(mod, "name")] = t.cast(Plugin, mod)  # type: ignore
        print(f"[plugins] ✅ loaded {modname} as '{getattr(mod, 'name')}'")
    return _REGISTRY


def get(name: str) -> Plugin:
    return _REGISTRY[name]


def list_plugins() -> list[str]:
    return list(_REGISTRY.keys())


if __name__ == "__main__":
    reg = load_all()
    names = list_plugins()
    print(f"Loaded plugins ({len(names)}): {names}")
    for n in names:
        mod = get(n)
        print(f"- {n}: {getattr(mod, 'description', '(no description)')}")
