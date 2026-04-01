"""Outlook for Mac OLM converter."""

from .extractor import extract
from .writer import write_all

__version__ = "0.1.0"

__all__ = ["extract", "write_all", "__version__"]
