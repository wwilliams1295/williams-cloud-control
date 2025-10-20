# core/capabilities.py
from dataclasses import dataclass, field
from typing import Callable, Any


@dataclass
class CapabilitySpec:
    name: str
    version: str
    description: str
    entrypoint: Callable[..., dict[str, Any]]
    inputs_schema: dict[str, Any]
    outputs_schema: dict[str, Any]
    example: dict[str, Any] = field(default_factory=dict)


class CapabilityRegistry:
    def __init__(self):
        self._caps: dict[str, CapabilitySpec] = {}

    def register(self, spec: CapabilitySpec):
        self._caps[spec.name] = spec

    def get(self, name: str) -> CapabilitySpec:
        return self._caps[name]

    def list_specs(self) -> list[CapabilitySpec]:
        return list(self._caps.values())


REGISTRY = CapabilityRegistry()


def describe_caps() -> list[dict[str, Any]]:
    """A compact, model-friendly catalog."""
    out = []
    for s in REGISTRY.list_specs():
        out.append(
            {
                "name": s.name,
                "desc": s.description,
                "inputs": list(s.inputs_schema.get("properties", {}).keys()),
                "example": s.example,
            }
        )
    return out
