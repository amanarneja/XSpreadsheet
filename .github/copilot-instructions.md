<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Excel MCP Server Project

This is an MCP (Model Context Protocol) Server project that provides Excel file manipulation capabilities.

## Key Components

- **Excel Library** (`src/excel_mcp_server/excel_library.py`): Core Excel manipulation functionality using openpyxl
- **MCP Server** (`src/excel_mcp_server/server.py`): FastMCP server that exposes Excel operations as tools

## Development Guidelines

- Use type hints throughout the codebase
- Follow Python best practices for error handling
- Log errors appropriately using the logging module
- Use openpyxl for Excel file manipulation
- Follow MCP protocol standards

## MCP Information

You can find more info and examples at https://modelcontextprotocol.io/llms-full.txt

## Available Tools

The server provides the following MCP tools:
- `read_excel_file` - Read data from Excel files
- `write_excel_file` - Write data to Excel files  
- `get_worksheet_info` - Get worksheet information
- `add_worksheet` - Add new worksheets
- `update_cell` - Update specific cells
- `apply_formula` - Apply Excel formulas
- `format_cells` - Apply cell formatting
- `create_chart` - Create charts and graphs

## Resources

The server provides these MCP resources:
- `excel://help` - Help documentation
- `excel://examples` - Usage examples
