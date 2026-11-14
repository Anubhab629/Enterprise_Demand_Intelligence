"""Semantic component registry for preventing duplicate dashboard content.

The registry is intentionally independent of Streamlit and dataset schemas.
Create one registry per dashboard render/session and share it across pages.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from typing import Any, Iterable


def _canonical(value: Any) -> Any:
    """Convert nested values into deterministic JSON-compatible structures."""
    if isinstance(value, dict):
        return {str(key): _canonical(value[key]) for key in sorted(value, key=str)}
    if isinstance(value, (list, tuple)):
        return [_canonical(item) for item in value]
    if isinstance(value, (set, frozenset)):
        return sorted((_canonical(item) for item in value), key=lambda item: json.dumps(item, sort_keys=True, default=str))
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _normalise_fields(fields: Iterable[Any] | Any | None) -> tuple[str, ...]:
    if fields is None:
        return ()
    if isinstance(fields, (str, bytes)):
        values = [fields]
    else:
        try:
            values = list(fields)
        except TypeError:
            values = [fields]
    return tuple(sorted(str(field) for field in values))


@dataclass(frozen=True)
class ComponentSignature:
    """Identity for a component based on purpose, source, fields, and operation."""
    component_type: str
    purpose: str
    source_id: str
    fields: tuple[str, ...] = ()
    operation: str = ""
    parameters_json: str = "{}"

    @classmethod
    def build(
        cls,
        component_type: str,
        purpose: str,
        source_id: str,
        fields: Iterable[Any] | Any | None = None,
        operation: str = "",
        parameters: dict[str, Any] | None = None,
    ) -> "ComponentSignature":
        parameters_json = json.dumps(
            _canonical(parameters or {}),
            sort_keys=True,
            ensure_ascii=False,
            separators=(",", ":"),
            default=str,
        )
        return cls(
            component_type=str(component_type).strip().casefold(),
            purpose=str(purpose).strip().casefold(),
            source_id=str(source_id).strip(),
            fields=_normalise_fields(fields),
            operation=str(operation).strip().casefold(),
            parameters_json=parameters_json,
        )

    @property
    def key(self) -> str:
        payload = json.dumps(
            {
                "component_type": self.component_type,
                "purpose": self.purpose,
                "source_id": self.source_id,
                "fields": self.fields,
                "operation": self.operation,
                "parameters": json.loads(self.parameters_json),
            },
            sort_keys=True,
            ensure_ascii=False,
            separators=(",", ":"),
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


class ComponentRegistry:
    """Track semantic component identities across multiple dashboard pages."""

    def __init__(self) -> None:
        self._components: dict[str, str] = {}

    def register(self, signature: ComponentSignature, page: str) -> bool:
        """Register a component; return False when the same semantic component exists.

        Reusing a signature on another page is considered a duplicate. The first
        registration remains authoritative.
        """
        if not isinstance(signature, ComponentSignature):
            raise TypeError("signature must be a ComponentSignature")
        page_name = str(page).strip() or "unknown"
        if signature.key in self._components:
            return False
        self._components[signature.key] = page_name
        return True

    def contains(self, signature: ComponentSignature) -> bool:
        return signature.key in self._components

    def page_for(self, signature: ComponentSignature) -> str | None:
        return self._components.get(signature.key)

    def clear(self) -> None:
        self._components.clear()

    def __len__(self) -> int:
        return len(self._components)

    def snapshot(self) -> dict[str, str]:
        """Return a defensive copy suitable for diagnostics."""
        return dict(self._components)
