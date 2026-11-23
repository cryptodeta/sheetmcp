"""Runtime configuration (env vars / CLI flags)."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

TRUE = {"1", "true", "yes", "on", "y"}


def _bool(name: str, default: bool) -> bool:
    v = os.environ.get(name)
    return default if v is None or v == "" else v.strip().lower() in TRUE


def _int(name: str, default: int) -> int:
    v = os.environ.get(name)
    try:
        return int(v) if v else default
    except ValueError:
        return default


@dataclass
class Config:
    """Engine configuration.

    roots:        directories the engine may touch (sandbox). Relative paths resolve against roots[0].
    read_only:    refuse every write.
    backup:       copy a file to ``.sheetmcp_backups/`` before overwriting it.
    max_backups:  backups kept per file.
    max_rows:     default row limit for outputs (protects LLM context).
    max_cell_chars: long strings are truncated in outputs to this length.
    """

    roots: list[str] = field(default_factory=lambda: [os.getcwd()])
    read_only: bool = False
    backup: bool = True
    max_backups: int = 10
    max_rows: int = 100
    max_cell_chars: int = 500

    def __post_init__(self) -> None:
        if not self.roots:
            self.roots = [os.getcwd()]
        self.roots = [str(Path(r).expanduser().resolve()) for r in self.roots]

    @classmethod
    def from_env(cls, **overrides: object) -> Config:
        roots_env = os.environ.get("SHEETMCP_ROOTS") or os.environ.get("SHEETMCP_ROOT")
        roots = [r for r in roots_env.split(os.pathsep) if r.strip()] if roots_env else [os.getcwd()]
        cfg = cls(
            roots=roots,
            read_only=_bool("SHEETMCP_READ_ONLY", False),
            backup=_bool("SHEETMCP_BACKUP", True),
            max_backups=_int("SHEETMCP_MAX_BACKUPS", 10),
            max_rows=_int("SHEETMCP_MAX_ROWS", 100),
            max_cell_chars=_int("SHEETMCP_MAX_CELL_CHARS", 500),
        )
        for k, v in overrides.items():
            if v is not None:
                setattr(cfg, k, v)
        cfg.__post_init__()
        return cfg
