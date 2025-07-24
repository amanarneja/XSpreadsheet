#!/usr/bin/env python3
"""
Setup script for Excel MCP Server
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the README file
readme_file = Path(__file__).parent / "README.md"
long_description = readme_file.read_text(encoding="utf-8") if readme_file.exists() else ""

setup(
    name="excel-mcp-server",
    version="1.0.0",
    description="A Model Context Protocol server for Excel file manipulation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Aman Arneja",
    author_email="aman.arneja@hotmail.com",
    url="https://github.com/amanarneja/excel-mcp-server",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.8",
    install_requires=[
        "mcp>=1.0.0",
        "openpyxl>=3.0.0",
        "pandas>=1.3.0",
        "anyio>=3.0.0",
        "xlsxwriter>=3.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=6.0.0",
            "pytest-asyncio>=0.18.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
            "mypy>=0.950",
        ],
    },
    entry_points={
        "console_scripts": [
            "excel-mcp-server=excel_mcp_server.server:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Office/Business :: Office Suites",
    ],
    keywords="excel mcp model-context-protocol spreadsheet automation",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/excel-mcp-server/issues",
        "Source": "https://github.com/yourusername/excel-mcp-server",
        "Documentation": "https://github.com/yourusername/excel-mcp-server#readme",
    },
)
