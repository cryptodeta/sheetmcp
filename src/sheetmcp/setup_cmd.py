"""`sheetmcp install|uninstall|doctor`: one-command client setup and self-diagnostics."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import platform
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

SERVER_NAME = "sheetmcp"
CLIENTS = ("claude-desktop", "cursor", "vscode", "windsurf", "claude-code")


# ---------------------------------------------------------------- config locations
def config_path(client: str, project: Path) -> Path:
    home = Path.home()
    system = platform.system()
    if client == "claude-desktop":
        if system == "Darwin":
            return home / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
        if system == "Windows":
            base = Path(os.environ.get("APPDATA") or home / "AppData" / "Roaming")
            return base / "Claude" / "claude_desktop_config.json"
        return (
            Path(os.environ.get("XDG_CONFIG_HOME") or home / ".config")
            / "Claude"
            / "claude_desktop_config.json"
        )
    if client == "cursor":
        return home / ".cursor" / "mcp.json"
    if client == "windsurf":
        return home / ".codeium" / "windsurf" / "mcp_config.json"
    if client == "vscode":
        return project / ".vscode" / "mcp.json"
    if client == "claude-code":
        return project / ".mcp.json"
    raise ValueError(client)


def detected(client: str, project: Path) -> bool:
    """A client counts as installed when its config folder exists (project-level clients: always offered)."""
    if client in ("vscode", "claude-code"):
        return False
    return config_path(client, project).parent.exists()


def server_entry(client: str, roots: list[str], read_only: bool, extra: list[str]) -> dict[str, Any]:
    args = ["-m", "sheetmcp", *roots]
    if read_only:
        args.append("--read-only")
    args += extra
    entry: dict[str, Any] = {"command": sys.executable, "args": args}
    if client == "vscode":
        entry = {"type": "stdio", **entry}
    return entry


def _key(client: str) -> str:
    return "servers" if client == "vscode" else "mcpServers"


def _load(p: Path) -> dict[str, Any]:
    if not p.exists() or not p.read_text(encoding="utf-8").strip():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise SystemExit(f"sheetmcp: {p} is not valid JSON ({e}); fix or move it, then retry.") from e
    if not isinstance(data, dict):
        raise SystemExit(f"sheetmcp: {p} does not contain a JSON object.")
    return data


def _save(p: Path, data: dict[str, Any]) -> str | None:
    p.parent.mkdir(parents=True, exist_ok=True)
    backup = None
    if p.exists():
        backup = str(p.with_name(p.name + f".bak-{dt.datetime.now():%Y%m%d-%H%M%S}"))
        shutil.copy2(p, backup)
    tmp = p.with_name(p.name + ".tmp")
    tmp.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, p)
    return backup


# ---------------------------------------------------------------- install / uninstall
def _targets(choice: list[str] | None, project: Path) -> list[str]:
    if not choice or choice == ["auto"]:
        found = [c for c in CLIENTS if detected(c, project)]
        return found
    if "all" in choice:
        return list(CLIENTS)
    bad = [c for c in choice if c not in CLIENTS]
    if bad:
        raise SystemExit(f"sheetmcp: unknown client(s) {bad}. Choose from: {', '.join(CLIENTS)}, all")
    return choice


def install(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(
        prog="sheetmcp install",
        description="Add SheetMCP to your AI apps (Claude Desktop, Cursor, VS Code, Windsurf, Claude Code).",
    )
    ap.add_argument("roots", nargs="*", help="Folder(s) the AI may access (default: the current folder).")
    ap.add_argument(
        "--client",
        "-c",
        action="append",
        help=f"Client to configure (repeatable): {', '.join(CLIENTS)}, all. Default: detected apps.",
    )
    ap.add_argument("--read-only", action="store_true", help="Install in read-only mode (no write tools).")
    ap.add_argument("--name", default=SERVER_NAME, help="Server name in the client config.")
    ap.add_argument("--project", default=".", help="Project folder for VS Code / Claude Code configs.")
    ap.add_argument("--dry-run", action="store_true", help="Show what would change without writing.")
    ap.add_argument("--print", dest="print_only", action="store_true", help="Only print the JSON snippet.")
    a = ap.parse_args(argv)
    roots = [str(Path(r).expanduser().resolve()) for r in (a.roots or ["."])]
    for r in roots:
        if not Path(r).is_dir():
            print(f"sheetmcp: folder not found: {r}", file=sys.stderr)
            return 2
    project = Path(a.project).expanduser().resolve()
    if a.print_only:
        print(json.dumps({"mcpServers": {a.name: server_entry("cursor", roots, a.read_only, [])}}, indent=2))
        return 0
    targets = _targets(a.client, project)
    if not targets:
        print(
            "No supported AI app was detected. Re-run with --client <name> (claude-desktop, cursor, vscode, "
            "windsurf, claude-code) or paste this into your client's MCP config:\n"
        )
        print(json.dumps({"mcpServers": {a.name: server_entry("cursor", roots, a.read_only, [])}}, indent=2))
        return 1
    for client in targets:
        p = config_path(client, project)
        data = _load(p)
        servers = data.setdefault(_key(client), {})
        if not isinstance(servers, dict):
            print(f"sheetmcp: '{_key(client)}' in {p} is not an object; skipped.", file=sys.stderr)
            continue
        action = "update" if a.name in servers else "add"
        servers[a.name] = server_entry(client, roots, a.read_only, [])
        if a.dry_run:
            print(f"[dry-run] would {action} '{a.name}' in {p}")
            continue
        backup = _save(p, data)
        print(
            f"{'Updated' if action == 'update' else 'Added'} '{a.name}' for {client}: {p}"
            + (f" (backup: {backup})" if backup else "")
        )
    if not a.dry_run:
        print(f"\nAllowed folder(s): {', '.join(roots)}{' (read-only)' if a.read_only else ''}")
        print(
            "Restart the app (Claude Desktop / Windsurf) or reload MCP servers (Cursor / VS Code) to start using it."
        )
        print('Try asking: "Take a quick look at my spreadsheets and tell me what\'s there."')
    return 0


def uninstall(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(
        prog="sheetmcp uninstall", description="Remove SheetMCP from client configs."
    )
    ap.add_argument("--client", "-c", action="append", help="Client(s); default: every client that has it.")
    ap.add_argument("--name", default=SERVER_NAME)
    ap.add_argument("--project", default=".")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)
    project = Path(a.project).expanduser().resolve()
    targets = list(CLIENTS) if not a.client or "all" in a.client else _targets(a.client, project)
    removed = 0
    for client in targets:
        p = config_path(client, project)
        if not p.exists():
            continue
        data = _load(p)
        servers = data.get(_key(client))
        if isinstance(servers, dict) and a.name in servers:
            removed += 1
            if a.dry_run:
                print(f"[dry-run] would remove '{a.name}' from {p}")
                continue
            del servers[a.name]
            _save(p, data)
            print(f"Removed '{a.name}' from {client}: {p}")
    if not removed:
        print(f"'{a.name}' was not found in any client config.")
    return 0


# ---------------------------------------------------------------- doctor
def doctor(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(
        prog="sheetmcp doctor", description="Check that SheetMCP works on this machine."
    )
    ap.add_argument(
        "roots", nargs="*", help="Folder(s) to check (default: $SHEETMCP_ROOTS or current folder)."
    )
    ap.add_argument("--project", default=".")
    ap.add_argument("--json", action="store_true", help="Machine-readable output.")
    a = ap.parse_args(argv)
    checks: list[dict[str, Any]] = []

    def add(name: str, ok: bool | None, detail: str) -> None:
        checks.append(
            {"check": name, "status": "ok" if ok else ("warn" if ok is None else "fail"), "detail": detail}
        )

    add("python", sys.version_info >= (3, 10), f"{platform.python_version()} ({sys.executable})")
    try:
        import duckdb
        import mcp
        import openpyxl

        from sheetmcp.core import Config, Sheets, __version__

        add(
            "packages",
            True,
            f"sheetmcp {__version__}, duckdb {duckdb.__version__}, openpyxl {openpyxl.__version__}, "
            f"mcp {getattr(mcp, '__version__', None) or _dist_version('mcp')}",
        )
    except Exception as e:  # pragma: no cover - broken install
        add("packages", False, f"import failed: {e}")
        return _report(checks, a.json)
    try:
        import xlrd  # noqa: F401

        add("xls support", True, "xlrd installed: legacy .xls files can be read")
    except ImportError:
        add("xls support", None, "optional: pip install 'sheetmcp[xls]' to read legacy .xls files")
    overrides: dict[str, Any] = {"roots": a.roots} if a.roots else {}
    try:
        cfg = Config.from_env(**overrides)
        S = Sheets(cfg)
        for r in S.sb.roots:
            w = os.access(r, os.W_OK)
            add(f"root {r}", None if not w else True, "readable and writable" if w else "read-only folder")
    except Exception as e:
        add("roots", False, str(e))
        return _report(checks, a.json)
    try:
        tmp = Path(tempfile.gettempdir())
        with tempfile.NamedTemporaryFile(dir=tmp):
            pass
        add("temp folder", True, str(tmp))
    except OSError as e:
        add("temp folder", False, f"{e}")
    with tempfile.TemporaryDirectory() as td:
        try:
            T = Sheets(Config(roots=[td]))
            T.write_range("t.xlsx", values=[["a", "b"], [1, 2], [3, 4]])
            T.write_range("t.xlsx", cells={"C2": "=A2+B2"})
            v = T.read_range("t.xlsx", range="C2")["values"][0][0]
            n = T.query("SELECT sum(a) AS s FROM t")["rows"][0][0]
            ok = v == 3 and n == 4
            add(
                "self-test", ok, "write xlsx, evaluate formula, SQL query" + ("" if ok else f": got {v}, {n}")
            )
        except Exception as e:
            add("self-test", False, f"{type(e).__name__}: {e}")
        try:
            T.query(f"SELECT * FROM read_csv('{Path.home() / 'x.csv'}')")
            add("sandbox", False, "DuckDB could read outside the allowed folder")
        except Exception:
            add("sandbox", True, "SQL cannot reach files outside the allowed folders")
    project = Path(a.project).resolve()
    found = []
    for c in CLIENTS:
        p = config_path(c, project)
        if p.exists():
            try:
                has = SERVER_NAME in (json.loads(p.read_text(encoding="utf-8") or "{}").get(_key(c)) or {})
            except (json.JSONDecodeError, AttributeError):
                has = False
            found.append(f"{c}: {'configured' if has else 'not configured'}")
    add(
        "clients",
        None if not any(f.split(": ")[1] == "configured" for f in found) else True,
        "; ".join(found) if found else "no client configs found; run `sheetmcp install --client <name>`",
    )
    return _report(checks, a.json)


def _dist_version(name: str) -> str:
    from importlib.metadata import PackageNotFoundError, version

    try:
        return version(name)
    except PackageNotFoundError:
        return "?"


def _report(checks: list[dict[str, Any]], as_json: bool) -> int:
    failed = any(c["status"] == "fail" for c in checks)
    if as_json:
        print(json.dumps({"ok": not failed, "checks": checks}, indent=2))
    else:
        icon = {"ok": "[ok]  ", "warn": "[warn]", "fail": "[FAIL]"}
        for c in checks:
            print(f"{icon[c['status']]} {c['check']}: {c['detail']}")
        print("\nAll good." if not failed else "\nSome checks failed; see above.")
    return 1 if failed else 0


COMMANDS = {"install": install, "init": install, "uninstall": uninstall, "doctor": doctor}
