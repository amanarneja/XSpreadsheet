# Contributing to Excel MCP Server

Thank you for your interest in contributing to the Excel MCP Server! We welcome contributions from the community.

## How to Contribute

### Reporting Issues

If you find a bug or have a feature request:

1. **Check existing issues** first to avoid duplicates
2. **Create a new issue** with a clear title and description
3. **Include steps to reproduce** for bugs
4. **Add relevant labels** to help categorize the issue

### Making Changes

1. **Fork the repository** to your GitHub account
2. **Create a new branch** for your feature or bug fix:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **Make your changes** following the coding standards below
4. **Test your changes** thoroughly
5. **Commit your changes** with descriptive commit messages
6. **Push to your fork** and submit a pull request

### Development Setup

1. Clone your fork:
   ```bash
   git clone https://github.com/YOUR-USERNAME/excel-mcp-server.git
   cd excel-mcp-server
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   source venv/bin/activate  # macOS/Linux
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run tests to ensure everything works:
   ```bash
   python test_server.py
   python verify.py
   ```

### Coding Standards

- **Python Style**: Follow PEP 8 guidelines
- **Type Hints**: Use type hints for all function parameters and return values
- **Documentation**: Add docstrings to all functions and classes
- **Error Handling**: Include proper error handling and logging
- **Testing**: Add tests for new functionality

### Testing

Before submitting a pull request:

1. **Run the test suite**:
   ```bash
   python test_server.py
   ```

2. **Test with the demo**:
   ```bash
   python demo.py
   ```

3. **Verify the MCP server**:
   ```bash
   python verify.py
   ```

4. **Test manual integration** with Claude Desktop if possible

### Pull Request Guidelines

- **Clear title** describing the change
- **Detailed description** of what was changed and why
- **Reference any related issues** using "Fixes #123" or "Relates to #123"
- **Include screenshots** for UI changes
- **Keep changes focused** - one feature or fix per PR

### Code Review Process

1. All pull requests require review before merging
2. Changes may be requested for code quality, testing, or documentation
3. Once approved, the PR will be merged by a maintainer

## Development Areas

We're particularly interested in contributions in these areas:

### High Priority
- **Additional chart types** (bubble, doughnut, radar charts)
- **Advanced formatting** (conditional formatting, data validation)
- **Performance optimizations** for large files
- **Error handling improvements**

### Medium Priority
- **CSV integration** (import/export functionality)
- **Template system** (pre-built Excel templates)
- **Data analysis tools** (pivot tables, statistical functions)
- **Batch operations** (process multiple files)

### Future Enhancements
- **Cloud storage integration** (Azure, AWS, Google Drive)
- **Real-time collaboration** features
- **Advanced security** (encryption, access controls)
- **Plugin system** for custom operations

## Questions?

If you have questions about contributing:

- **Open an issue** for general questions
- **Check the README** for basic usage information
- **Review existing code** to understand patterns and conventions

Thank you for contributing to Excel MCP Server!
