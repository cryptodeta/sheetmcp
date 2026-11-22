from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

from sheetmcp_core import Config, Sheets

ROOT = Path(__file__).resolve().parent.parent
_spec = importlib.util.spec_from_file_location("make_samples", ROOT / "examples" / "make_samples.py")
make_samples = importlib.util.module_from_spec(_spec)
sys.modules["make_samples"] = make_samples
_spec.loader.exec_module(make_samples)


@pytest.fixture()
def data_dir(tmp_path: Path) -> Path:
    """A fresh copy of the sample files (xlsx with formulas, csv, parquet, json) per test."""
    return make_samples.make(tmp_path / "data")


@pytest.fixture()
def sheets(data_dir: Path) -> Sheets:
    return Sheets(Config(roots=[str(data_dir)]))


@pytest.fixture()
def ro_sheets(data_dir: Path) -> Sheets:
    return Sheets(Config(roots=[str(data_dir)], read_only=True))
