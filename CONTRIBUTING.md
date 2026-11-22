# Contributing

Thanks for helping! The bar: **it must work reliably, and errors must say how to fix them.**

## Setup

```bash
git clone <repo> && cd sheetmcp
uv sync                      # installs all workspace packages (editable) + dev tools
uv run pytest                # run tests
uv run ruff check . && uv run ruff format .
```

## Where things go

* Engine logic → `packages/core/src/sheetmcp_core/` (no MCP imports there).
* New MCP tool → add the engine method first (with tests), then a thin wrapper in
  `packages/server/src/sheetmcp/server.py` with `Annotated[..., Field(description=...)]` parameters.
  Run `uv run python scripts/gen_tool_docs.py` to refresh `docs/tools.md`.
* Prefer extending an existing tool (an `action` or option) over adding a new one; we keep the tool count small.
* New formula function → `evaluator.py` (`@fn("NAME")`) + a case in `tests/test_evaluator.py`.

## Rules

* Raise `SheetError(message, hint)` for anything a user/LLM can fix; the hint must be actionable.
* Every output that can grow must be limited/paged.
* Writes go through `save_workbook` / `atomic_write` (backup + atomic replace) and respect `read_only` and `dry_run`.
* Keep dependencies minimal; core depends only on `openpyxl` and `duckdb`.
* Tests generate their own fixture files (see `examples/make_samples.py`); do not commit binary fixtures to `tests/`.

## Releasing

1. Bump versions in the three `packages/*/pyproject.toml`, `sheetmcp_core/__init__.py` and the pinned
   `sheetmcp-core==` dependency; update `CHANGELOG.md`.
2. `uv build --all-packages` then `uv publish` (core first).
