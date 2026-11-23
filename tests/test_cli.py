import json
from pathlib import Path

from sheetmcp.cli.main import main


def test_cli_query_and_convert(data_dir: Path, capsys):
    assert main(["--root", str(data_dir), "query", "SELECT count(*) AS n FROM customers"]) == 0
    assert json.loads(capsys.readouterr().out)["rows"] == [[13]]
    assert main(["--root", str(data_dir), "convert", "sales.xlsx", "sales.parquet"]) == 0
    assert (data_dir / "sales.parquet").exists()


def test_cli_read_markdown_and_errors(data_dir: Path, capsys):
    assert main(["--root", str(data_dir), "read", "sales.xlsx", "-s", "Summary", "-r", "A1:B3", "--md"]) == 0
    assert "| 1 | Region | Revenue |" in capsys.readouterr().out
    assert main(["--root", str(data_dir), "describe", "missing.xlsx"]) == 1
    assert "not found" in capsys.readouterr().err


def test_cli_write(data_dir: Path, capsys):
    assert main(["--root", str(data_dir), "write", "sales.xlsx", "-s", "Summary", "G1=42", "G2==G1*2"]) == 0
    out = json.loads(capsys.readouterr().out)
    assert out["cells_written"] == 2
