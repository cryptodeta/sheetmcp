"""Errors with actionable hints. Every error the engine raises is a SheetError."""

from __future__ import annotations


class SheetError(Exception):
    """An error meant to be shown to a human or an LLM.

    ``message`` says what went wrong, ``hint`` says how to fix it.
    """

    def __init__(self, message: str, hint: str | None = None) -> None:
        self.message = message
        self.hint = hint
        super().__init__(self.__str__())

    def __str__(self) -> str:
        return f"{self.message} Hint: {self.hint}" if self.hint else self.message


class ReadOnlyError(SheetError):
    def __init__(self, action: str) -> None:
        super().__init__(
            f"Cannot {action}: the server is running in read-only mode.",
            "Restart without --read-only (or unset SHEETMCP_READ_ONLY) to allow writes.",
        )
