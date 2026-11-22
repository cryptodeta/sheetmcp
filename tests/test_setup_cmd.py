"""`sheetmcp install / uninstall / doctor`."""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

from sheetmcp import setup_cmd
from sheetmcp.__main__ import main


@pytest.fixture()
def home(tmp_path: Path, monkeypatch) -> Path:
    h = tmp_path / "home"
    h.mkdir()
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: h))
    monkeypatch.setenv("HOME", str(h))
    monkeypatch.setenv("APPDATA", str(h / "AppData"))
    monkeypatch.delenv("XDG_CONFIG_HOME", raising=False)
    return h


def _run(argv: list[str]) -> int:
    with pytest.raises(SystemExit) as e:
        main(argv)
    return e.value.code


def test_install_detected_and_merge(home: Path, data_dir: Path, capsys):
    cur = home / ".cursor" / "mcp.json"
    cur.parent.mkdir()
    cur.write_text(json.dumps({"mcpServers": {"other": {"command": "x"}}}), encoding="utf-8")
    assert _run(["install", str(data_dir)]) == 0
    cfg = json.loads(cur.read_text(encoding="utf-8"))
    assert cfg["mcpServers"]["other"] == {"command": "x"}
    entry = cfg["mcpServers"]["sheetmcp"]
    assert entry["command"] == sys.executable and entry["args"] == ["-m", "sheetmcp", str(data_dir.resolve())]
    assert list(cur.parent.glob("mcp.json.bak-*")), "existing config must be backed up"
    assert "Added 'sheetmcp' for cursor" in capsys.readouterr().out
    assert _run(["install", str(data_dir), "--read-only"]) == 0
    assert "--read-only" in json.loads(cur.read_text(encoding="utf-8"))["mcpServers"]["sheetmcp"]["args"]


def test_install_all_clients_paths(home: Path, data_dir: Path, tmp_path: Path, monkeypatch):
    proj = tmp_path / "proj"
    proj.mkdir()
    for system in ("Linux", "Darwin", "Windows"):
        monkeypatch.setattr(setup_cmd.platform, "system", lambda s=system: s)
        assert _run(["install", str(data_dir), "--client", "all", "--project", str(proj)]) == 0
        claude = setup_cmd.config_path("claude-desktop", proj)
        assert json.loads(claude.read_text(encoding="utf-8"))["mcpServers"]["sheetmcp"]
    assert (home / ".config" / "Claude" / "claude_desktop_config.json").exists()
    assert (home / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json").exists()
    assert (home / "AppData" / "Claude" / "claude_desktop_config.json").exists()
    vs = json.loads((proj / ".vscode" / "mcp.json").read_text(encoding="utf-8"))
    assert vs["servers"]["sheetmcp"]["type"] == "stdio"
    assert json.loads((proj / ".mcp.json").read_text(encoding="utf-8"))["mcpServers"]["sheetmcp"]
    assert (home / ".codeium" / "windsurf" / "mcp_config.json").exists()
    assert _run(["uninstall", "--project", str(proj)]) == 0
    assert "sheetmcp" not in json.loads((proj / ".mcp.json").read_text(encoding="utf-8"))["mcpServers"]


def test_install_dry_run_print_and_errors(home: Path, data_dir: Path, capsys):
    assert _run(["install", str(data_dir), "--client", "cursor", "--dry-run"]) == 0
    assert "[dry-run]" in capsys.readouterr().out and not (home / ".cursor").exists()
    assert _run(["install", str(data_dir), "--print"]) == 0
    assert json.loads(capsys.readouterr().out)["mcpServers"]["sheetmcp"]["command"] == sys.executable
    assert _run(["install", str(data_dir)]) == 1  # nothing detected -> prints snippet
    assert "mcpServers" in capsys.readouterr().out
    assert _run(["install", str(data_dir / "missing")]) == 2
    with pytest.raises(SystemExit):
        main(["install", str(data_dir), "--client", "notepad"])
    bad = home / ".cursor" / "mcp.json"
    bad.parent.mkdir()
    bad.write_text("{oops", encoding="utf-8")
    with pytest.raises(SystemExit, match="not valid JSON"):
        main(["install", str(data_dir), "--client", "cursor"])


def test_doctor(home: Path, data_dir: Path, capsys):
    assert _run(["doctor", str(data_dir)]) == 0
    out = capsys.readouterr().out
    assert "[ok]   self-test" in out and "[ok]   sandbox" in out and "All good." in out
    assert _run(["doctor", str(data_dir), "--json"]) == 0
    rep = json.loads(capsys.readouterr().out)
    assert rep["ok"] and {c["check"] for c in rep["checks"]} >= {"python", "packages", "self-test", "sandbox"}
    assert _run(["doctor", str(data_dir / "nope")]) == 1
