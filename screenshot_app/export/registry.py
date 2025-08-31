# export/registry.py
from __future__ import annotations
from typing import Dict, Type
from .base import BaseExporter

_REGISTRY: Dict[str, Type[BaseExporter]] = {}

def register(cls: Type[BaseExporter]):
    if not issubclass(cls, BaseExporter):
        raise TypeError("Exporter subclass required (BaseExporter)")
    if not getattr(cls, "name", None):
        raise ValueError("Exporter must define .name")
    _REGISTRY[cls.name] = cls
    return cls

def get(name: str) -> BaseExporter:
    if not _REGISTRY:
        raise RuntimeError("No exporters are registered. Did you import export package?")
    if name not in _REGISTRY:
        raise KeyError(f"Exporter '{name}' is not registered. Available: {list(_REGISTRY.keys())}")
    return _REGISTRY[name]()

def all_names():
    return list(_REGISTRY.keys())

def get_default_exporter_name() -> str:
    if not _REGISTRY:
        raise RuntimeError("No exporters are registered. Did you import export package?")
    return "excel" if "excel" in _REGISTRY else next(iter(_REGISTRY.keys()))
