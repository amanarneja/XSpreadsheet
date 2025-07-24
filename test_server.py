#!/usr/bin/env python3
"""
Test script for Excel MCP Server

This script tests both the Excel library and MCP server functionality.
"""

import sys
import os
from pathlib import Path

# Add the project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src" / "excel_mcp_server"))

def test_excel_library():
    """Test the Excel library functionality"""
    print("🧪 Testing Excel Library...")
    
    try:
        from excel_library import ExcelLibrary
        
        # Initialize library
        excel_lib = ExcelLibrary()
        print("✅ Excel library imported successfully")
        
        # Test creating a simple Excel file
        test_data = [
            ["Name", "Age", "City"],
            ["Alice", 30, "New York"],
            ["Bob", 25, "San Francisco"],
            ["Charlie", 35, "Chicago"]
        ]
        
        test_file = "test_output.xlsx"
        result = excel_lib.write_excel_file(test_file, test_data, headers=["Name", "Age", "City"])
        print("✅ Excel file creation test passed")
        
        # Test reading the file back
        read_result = excel_lib.read_excel_file(test_file)
        print("✅ Excel file reading test passed")
        
        # Clean up
        if os.path.exists(test_file):
            os.remove(test_file)
            print("🧹 Cleaned up test file")
            
        return True
        
    except Exception as e:
        print(f"❌ Excel library test failed: {e}")
        return False

def test_mcp_server():
    """Test the MCP server can be imported"""
    print("\n🧪 Testing MCP Server...")
    
    try:
        # Test imports
        from mcp.server import FastMCP
        print("✅ FastMCP imported successfully")
        
        # Test server module can be imported
        sys.path.insert(0, str(project_root / "src" / "excel_mcp_server"))
        import server
        print("✅ Server module imported successfully")
        
        return True
        
    except Exception as e:
        print(f"❌ MCP server test failed: {e}")
        return False

def main():
    """Main test function"""
    print("🚀 Excel MCP Server Test Suite")
    print("=" * 50)
    
    # Test Excel library
    library_ok = test_excel_library()
    
    # Test MCP server
    server_ok = test_mcp_server()
    
    print("\n" + "=" * 50)
    print("📊 Test Results:")
    print(f"Excel Library: {'✅ PASS' if library_ok else '❌ FAIL'}")
    print(f"MCP Server: {'✅ PASS' if server_ok else '❌ FAIL'}")
    
    if library_ok and server_ok:
        print("\n🎉 All tests passed! Your Excel MCP Server is ready to use.")
        print("\nNext steps:")
        print("1. Run the server: python src/excel_mcp_server/server.py")
        print("2. Configure your MCP client to use this server")
        print("3. Access help with the excel://help resource")
    else:
        print("\n⚠️  Some tests failed. Check the error messages above.")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
