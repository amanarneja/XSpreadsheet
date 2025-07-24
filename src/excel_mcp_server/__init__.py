"""
Excel MCP Server Package

This package provides an MCP server for Excel file manipulation with a comprehensive
library for reading, writing, and manipulating Excel files.
"""

from .server import main, mcp
from .excel_library import ExcelLibrary

__version__ = "1.0.0"
__author__ = "Excel MCP Server"

__all__ = ["main", "mcp", "ExcelLibrary"]
