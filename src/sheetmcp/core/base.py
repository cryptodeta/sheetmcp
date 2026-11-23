"""Shared state for all operation mixins."""

from __future__ import annotations

from typing import Any

from .config import Config
from .errors import ReadOnlyError
from .paths import Sandbox
from .util import jsonable

BIG_FILE_BYTES = 8 * 1024 * 1024  # above this, xlsx files are streamed in read-only mode


class Base:
    def __init__(self, config: Config | None = None) -> None:
        self.cfg = config or Config.from_env()
        self.sb = Sandbox(self.cfg.roots)

    def _limit(self, limit: int | None, hard_max: int = 10_000) -> int:
        n = self.cfg.max_rows if limit is None or limit <= 0 else int(limit)
        return max(1, min(n, hard_max))

    def _writable(self, action: str) -> None:
        if self.cfg.read_only:
            raise ReadOnlyError(action)

    def _j(self, v: Any) -> Any:
        return jsonable(v, self.cfg.max_cell_chars)
