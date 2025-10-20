# plugins/__init__.py
from typing import Protocol, Dict, Any, List, Optional, runtime_checkable


@runtime_checkable
class Plugin(Protocol):
    name: str
    description: str
    permissions: Optional[list[str]]

    def run(self, **kwargs) -> dict[str, Any]: ...
