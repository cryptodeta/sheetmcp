"""sheetmcp.core: fast, lightweight spreadsheet engine (xlsx / csv / parquet / json). No MCP dependency."""

from .. import __version__
from .api import Sheets
from .config import Config
from .errors import ReadOnlyError, SheetError
from .evaluator import Evaluator

__all__ = ["Config", "Evaluator", "ReadOnlyError", "SheetError", "Sheets", "__version__"]
