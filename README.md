# Excel MCP Server

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-compatible-green.svg)](https://modelcontextprotocol.io/)

A comprehensive Model Context Protocol (MCP) server that provides Excel file manipulation capabilities. This project includes both a standalone Excel library and an MCP server that exposes Excel operations as tools for AI assistants like Claude.

## ✨ Features

### 📊 Excel Operations
- **Read Excel Files**: Extract data from .xlsx and .xls files with range and sheet selection
- **Write Excel Files**: Create new Excel files with structured data and headers
- **Worksheet Management**: Add, remove, and manage multiple worksheets
- **Cell Operations**: Update individual cells and ranges with values or formulas
- **Formula Support**: Apply and calculate Excel formulas including cross-sheet references
- **Formatting**: Apply styling, fonts, colors, and advanced formatting options
- **Charts**: Create various chart types (line, bar, pie, scatter) with customization

### 🔧 MCP Integration
- **FastMCP Server**: High-performance MCP server implementation using the official MCP SDK
- **Resource Endpoints**: Built-in help documentation and usage examples
- **Tool Definitions**: All Excel operations exposed as MCP tools with full type hints
- **Error Handling**: Comprehensive error reporting and logging throughout

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/amanarneja/XSpreadsheet.git
   cd XSpreadsheet
   ```

2. **Create a virtual environment:**
   ```bash
   python -m venv venv
   ```

3. **Activate the virtual environment:**
   ```bash
   # Windows
   venv\Scripts\activate
   
   # macOS/Linux
   source venv/bin/activate
   ```

4. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the MCP Server

```bash
cd src/excel_mcp_server
python server.py
```

The server will start and listen for MCP protocol messages on standard input/output.

### Using the Excel Library

```python
from src.excel_mcp_server.excel_library import ExcelLibrary

# Initialize the library
excel_lib = ExcelLibrary()

# Read an Excel file
data = excel_lib.read_excel_file("example.xlsx", "Sheet1")

# Write data to Excel
excel_lib.write_excel_file("output.xlsx", [
    ["Name", "Age", "City"],
    ["John", 30, "New York"],
    ["Jane", 25, "Boston"]
], headers=["Name", "Age", "City"])
```

## Available MCP Tools

### Core Operations

- **`read_excel_file`**
  - Read data from Excel files
  - Parameters: `file_path`, `sheet_name` (optional), `range` (optional)

- **`write_excel_file`**
  - Write data to Excel files
  - Parameters: `file_path`, `data`, `sheet_name` (optional), `headers` (optional)

- **`get_worksheet_info`**
  - Get information about worksheets
  - Parameters: `file_path`

### Worksheet Management

- **`add_worksheet`**
  - Add new worksheets
  - Parameters: `file_path`, `sheet_name`

### Cell Operations

- **`update_cell`**
  - Update specific cells
  - Parameters: `file_path`, `sheet_name`, `cell`, `value`

- **`apply_formula`**
  - Apply Excel formulas
  - Parameters: `file_path`, `sheet_name`, `cell`, `formula`

### Formatting and Charts

- **`format_cells`**
  - Apply cell formatting
  - Parameters: `file_path`, `sheet_name`, `range`, `format_options`

- **`create_chart`**
  - Create charts and graphs
  - Parameters: `file_path`, `sheet_name`, `data_range`, `chart_type`, `title` (optional), `position` (optional)

## MCP Resources

- **`excel://help`** - Comprehensive help documentation
- **`excel://examples`** - Usage examples and code snippets

## Configuration

### VS Code Integration

The project includes VS Code configuration for MCP debugging:

```json
{
  "servers": {
    "excel-mcp-server": {
      "type": "stdio",
      "command": "python",
      "args": ["src/excel_mcp_server/server.py"]
    }
  }
}
```

### Claude Desktop Integration

To use with Claude Desktop, add this to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "python",
      "args": ["path/to/XSpreadsheet/src/excel_mcp_server/server.py"],
      "cwd": "path/to/XSpreadsheet"
    }
  }
}
```

**Windows example:**
```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "python",
      "args": ["C:\\projects\\XSpreadsheet\\src\\excel_mcp_server\\server.py"],
      "cwd": "C:\\projects\\XSpreadsheet"
    }
  }
}
```

### Alternative: Module Execution
```json
{
  "mcpServers": {
    "excel-mcp-server": {
      "command": "python",
      "args": ["-m", "src.excel_mcp_server.server"],
      "cwd": "path/to/XSpreadsheet"
    }
  }
}
```

## Project Structure

```
Xpreadsheet/
├── src/
│   └── excel_mcp_server/
│       ├── __init__.py
│       ├── server.py           # MCP Server implementation
│       └── excel_library.py    # Core Excel functionality
├── .vscode/
│   └── mcp.json               # VS Code MCP configuration
├── .github/
│   └── copilot-instructions.md # Copilot development guidelines
├── requirements.txt           # Python dependencies
└── README.md                 # This file
```

## Dependencies

- **mcp**: Model Context Protocol SDK
- **openpyxl**: Excel file manipulation
- **pandas**: Data processing and analysis
- **anyio**: Async I/O support

## Development

### Adding New Features

1. **Excel Library**: Add new methods to `excel_library.py`
2. **MCP Tools**: Expose new methods as MCP tools in `server.py`
3. **Documentation**: Update help resources and README

## 🧪 Testing & Demo

### Run the Demo
```bash
python demo.py
```
This creates sample Excel files demonstrating all features.

### Verify Installation
```bash
python verify.py
```

### Test Individual Components
```bash
# Test the Excel library
python -c "from src.excel_mcp_server.excel_library import ExcelLibrary; print('✅ Library works!')"

# Test the MCP server (press Ctrl+C to stop)
python src/excel_mcp_server/server.py
```

## 📖 Usage Examples

### Reading Excel Data with Claude

Once configured, you can ask Claude:
> "Read the data from my sales.xlsx file, sheet called 'Q1 Data', range A1:D10"

Claude will use the `read_excel_file` tool:
```json
{
  "file_path": "sales.xlsx",
  "sheet_name": "Q1 Data", 
  "range": "A1:D10"
}
```

### Creating Charts with Claude

> "Create a line chart from my data.xlsx file showing the sales trend. Use data from A1:B12 on the 'Sales' sheet"

Claude will use the `create_chart` tool:
```json
{
  "file_path": "data.xlsx",
  "sheet_name": "Sales",
  "data_range": "A1:B12",
  "chart_type": "line",
  "title": "Sales Trend"
}
```

### Writing Excel Files

> "Create a new Excel file called 'report.xlsx' with this data: [['Name', 'Score'], ['Alice', 95], ['Bob', 87]]"

Claude will use the `write_excel_file` tool:
```json
{
  "file_path": "report.xlsx",
  "data": [["Alice", 95], ["Bob", 87]],
  "headers": ["Name", "Score"]
}
```

## 🤝 Contributing

We welcome contributions! Here's how to get started:

1. **Fork the repository** on GitHub
2. **Clone your fork** locally:
   ```bash
   git clone https://github.com/amanarneja/XSpreadsheet.git
   ```
3. **Create a feature branch**:
   ```bash
   git checkout -b feature/amazing-feature
   ```
4. **Make your changes** and add tests
5. **Run the verification**:
   ```bash
   python verify.py
   python demo.py
   ```
6. **Commit your changes**:
   ```bash
   git commit -m "Add amazing feature"
   ```
7. **Push to your branch**:
   ```bash
   git push origin feature/amazing-feature
   ```
8. **Open a Pull Request** on GitHub

### Development Guidelines
- Follow the existing code style and patterns
- Add type hints to all functions
- Include docstrings for new methods
- Update tests and documentation
- Ensure all demo scripts still work

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support & Resources

- **Issues**: [GitHub Issues](https://github.com/amanarneja/XSpreadsheet/issues)
- **MCP Documentation**: [Model Context Protocol](https://modelcontextprotocol.io/)
- **Claude Desktop**: [Anthropic Claude](https://claude.ai/)
- **Excel Library**: Built with [openpyxl](https://openpyxl.readthedocs.io/)

## 🌟 Acknowledgments

- Built using the [Model Context Protocol (MCP)](https://modelcontextprotocol.io/)
- Excel manipulation powered by [openpyxl](https://openpyxl.readthedocs.io/)
- Data processing with [pandas](https://pandas.pydata.org/)

---

**🚀 Ready to supercharge your Excel workflows with AI? Get started today!** 

[![GitHub stars](https://img.shields.io/github/stars/amanarneja/XSpreadsheet?style=social)](https://github.com/amanarneja/XSpreadsheet)
[![GitHub forks](https://img.shields.io/github/forks/amanarneja/XSpreadsheet?style=social)](https://github.com/amanarneja/XSpreadsheet)
