#!/usr/bin/env python3
"""
Excel MCP Server

An MCP server that provides Excel file manipulation capabilities including:
- Reading Excel files
- Writing Excel files
- Managing worksheets
- Cell operations
- Formula handling
- Data formatting
"""

import logging
from typing import Any, Dict, List, Optional, Union
import json
import os
from pathlib import Path

from mcp.server import FastMCP

# Handle both script and module execution
try:
    from .excel_library import ExcelLibrary
except ImportError:
    # Running as script, use absolute import
    from excel_library import ExcelLibrary

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("excel-mcp-server")

# Create the FastMCP server
mcp = FastMCP("Excel MCP Server")

# Initialize the Excel library
excel_lib = ExcelLibrary()

@mcp.resource("excel://help")
def get_help() -> str:
    """Get help documentation for Excel MCP Server"""
    return """
Excel MCP Server Help

This MCP server provides comprehensive Excel file manipulation capabilities.

Available Tools:
1. read_excel_file - Read data from Excel files
2. write_excel_file - Write data to Excel files
3. get_worksheet_info - Get information about worksheets
4. add_worksheet - Add new worksheets
5. update_cell - Update specific cells
6. apply_formula - Apply Excel formulas
7. format_cells - Apply formatting to cells
8. create_chart - Create charts and graphs

Features:
- Support for .xlsx and .xls files
- Multiple worksheet management
- Cell formatting and styling
- Formula calculations
- Chart creation
- Data import/export

For detailed usage examples, access the excel://examples resource.
"""

@mcp.resource("excel://examples")
def get_examples() -> str:
    """Get example usage content"""
    return """
Excel MCP Server Examples

1. Reading an Excel file:
   Tool: read_excel_file
   Arguments: {"file_path": "data.xlsx", "sheet_name": "Sheet1", "range": "A1:C10"}

2. Writing data to Excel:
   Tool: write_excel_file
   Arguments: {
     "file_path": "output.xlsx",
     "data": [["Name", "Age", "City"], ["John", 30, "New York"], ["Jane", 25, "Boston"]],
     "headers": ["Name", "Age", "City"]
   }

3. Adding a new worksheet:
   Tool: add_worksheet
   Arguments: {"file_path": "workbook.xlsx", "sheet_name": "NewSheet"}

4. Updating a cell:
   Tool: update_cell
   Arguments: {"file_path": "data.xlsx", "sheet_name": "Sheet1", "cell": "A1", "value": "Updated Value"}

5. Applying a formula:
   Tool: apply_formula
   Arguments: {"file_path": "calc.xlsx", "sheet_name": "Sheet1", "cell": "C1", "formula": "=A1+B1"}
   
   Cross-sheet reference example:
   Arguments: {"file_path": "calc.xlsx", "sheet_name": "Summary", "cell": "A1", "formula": "=SUM(Data!D:D)"}

6. Formatting cells:
   Tool: format_cells
   Arguments: {
     "file_path": "styled.xlsx",
     "sheet_name": "Sheet1",
     "range": "A1:C1",
     "format_options": {"bold": true, "font_size": 14, "background_color": "yellow"}
   }

7. Creating a chart:
   Tool: create_chart
   Arguments: {
     "file_path": "charts.xlsx",
     "sheet_name": "Data",
     "data_range": "A1:B10",
     "chart_type": "line",
     "title": "Sales Trend"
   }
"""

@mcp.tool()
def read_excel_file(file_path: str, sheet_name: Optional[str] = None, range: Optional[str] = None) -> Dict[str, Any]:
    """Read data from an Excel file"""
    try:
        return excel_lib.read_excel_file(file_path, sheet_name, range)
    except Exception as e:
        logger.error(f"Error reading Excel file: {str(e)}")
        raise

@mcp.tool()
def write_excel_file(file_path: str, data: List[List[Any]], sheet_name: Optional[str] = None, headers: Optional[List[str]] = None) -> Dict[str, Any]:
    """Write data to an Excel file"""
    try:
        return excel_lib.write_excel_file(file_path, data, sheet_name, headers)
    except Exception as e:
        logger.error(f"Error writing Excel file: {str(e)}")
        raise

@mcp.tool()
def get_worksheet_info(file_path: str) -> Dict[str, Any]:
    """Get information about worksheets in an Excel file"""
    try:
        return excel_lib.get_worksheet_info(file_path)
    except Exception as e:
        logger.error(f"Error getting worksheet info: {str(e)}")
        raise

@mcp.tool()
def add_worksheet(file_path: str, sheet_name: str) -> Dict[str, Any]:
    """Add a new worksheet to an existing Excel file"""
    try:
        return excel_lib.add_worksheet(file_path, sheet_name)
    except Exception as e:
        logger.error(f"Error adding worksheet: {str(e)}")
        raise

@mcp.tool()
def update_cell(file_path: str, sheet_name: str, cell: str, value: Any) -> Dict[str, Any]:
    """Update a specific cell in an Excel file"""
    try:
        return excel_lib.update_cell(file_path, sheet_name, cell, value)
    except Exception as e:
        logger.error(f"Error updating cell: {str(e)}")
        raise

@mcp.tool()
def apply_formula(file_path: str, sheet_name: str, cell: str, formula: str) -> Dict[str, Any]:
    """Apply a formula to a cell in an Excel file"""
    try:
        return excel_lib.apply_formula(file_path, sheet_name, cell, formula)
    except Exception as e:
        logger.error(f"Error applying formula: {str(e)}")
        raise

@mcp.tool()
def format_cells(file_path: str, sheet_name: str, range: str, format_options: Dict[str, Any]) -> Dict[str, Any]:
    """Apply formatting to cells in an Excel file"""
    try:
        return excel_lib.format_cells(file_path, sheet_name, range, format_options)
    except Exception as e:
        logger.error(f"Error formatting cells: {str(e)}")
        raise

@mcp.tool()
def create_chart(file_path: str, sheet_name: str, data_range: str, chart_type: str, title: Optional[str] = None, position: Optional[str] = None) -> Dict[str, Any]:
    """Create a chart in an Excel file"""
    try:
        return excel_lib.create_chart(file_path, sheet_name, data_range, chart_type, title, position)
    except Exception as e:
        logger.error(f"Error creating chart: {str(e)}")
        raise

def main():
    """Main entry point"""
    mcp.run()

if __name__ == "__main__":
    main()
