"""Sandbox: every path the engine touches must live inside an allowed root."""

from __future__ import annotations

import difflib
from pathlib import Path

from .errors import SheetError


class Sandbox:
    def __init__(self, roots: list[str]) -> None:
        self.roots = [Path(r).expanduser().resolve() for r in roots]
        for r in self.roots:
            if not r.is_dir():
                raise SheetError(
                    f"Allowed root '{r}' does not exist or is not a directory.",
                    "Pass an existing folder with --root or SHEETMCP_ROOTS.",
                )

    def inside(self, p: Path) -> bool:
        return any(p == r or r in p.parents for r in self.roots)

    def resolve(self, path: str, must_exist: bool = True, want: str = "file") -> Path:
        if path is None or str(path).strip() == "":
            raise SheetError(
                "A path is required.",
                f"Use a path relative to {self.roots[0]} or an absolute path inside it.",
            )
        raw = Path(str(path).strip()).expanduser()
        p = raw if raw.is_absolute() else self.roots[0] / raw
        p = p.resolve()
        if not self.inside(p):
            raise SheetError(
                f"Path '{path}' is outside the allowed directories.",
                "Allowed roots: " + ", ".join(str(r) for r in self.roots) + ". Use a path inside one of them "
                "or restart the server with an extra --root.",
            )
        if must_exist and not p.exists():
            hint = "Call list_files to see available files."
            parent = p.parent
            if parent.is_dir():
                names = [c.name for c in parent.iterdir()]
                close = difflib.get_close_matches(p.name, names, n=3, cutoff=0.5)
                if close:
                    hint = "Did you mean: " + ", ".join(close) + "? " + hint
            raise SheetError(f"'{path}' was not found.", hint)
        if must_exist and want == "file" and p.is_dir():
            raise SheetError(
                f"'{path}' is a folder, not a file.", "Pass a file path, or use list_files for folders."
            )
        if must_exist and want == "dir" and not p.is_dir():
            raise SheetError(f"'{path}' is not a folder.", "Pass a folder path.")
        return p

    def display(self, p: Path) -> str:
        """Short path for output: relative to its root when possible."""
        for r in self.roots:
            if p == r or r in p.parents:
                rel = p.relative_to(r)
                return str(rel) if len(self.roots) == 1 else str(p)
        return str(p)
