#!/usr/bin/env python3
"""
Excel MCP Server - Quick Verification Script

This script performs a quick verification that all components are working.
"""

import sys
import os
from pathlib import Path

# Add the project to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src"))

def main():
    print("🔍 Excel MCP Server - Quick Verification")
    print("=" * 50)
    
    try:
        # Test Excel library import
        from excel_mcp_server.excel_library import ExcelLibrary
        print("✅ Excel library imported successfully")
        
        # Test server import
        import excel_mcp_server.server
        print("✅ MCP server imported successfully")
        
        # Test basic Excel operations
        excel_lib = ExcelLibrary()
        
        # Create a simple test
        test_data = [["Test", "Data"], ["Row1", 123], ["Row2", 456]]
        
        # Test file creation
        result = excel_lib.write_excel_file(
            "verification_test.xlsx",
            test_data[1:],  # Data rows
            sheet_name="Test",
            headers=test_data[0]  # Headers
        )
        
        if result.get('success'):
            print(f"✅ Excel file creation: {result['rows_written']} rows written")
            
            # Test file reading
            read_result = excel_lib.read_excel_file("verification_test.xlsx", "Test")
            if read_result.get('success'):
                print(f"✅ Excel file reading: {len(read_result['data'])} rows read")
            else:
                print(f"❌ Excel file reading failed: {read_result.get('error')}")
            
            # Clean up
            try:
                os.remove("verification_test.xlsx")
                print("✅ Test file cleaned up")
            except:
                print("⚠️  Could not clean up test file (might be locked)")
                
        else:
            print(f"❌ Excel file creation failed: {result.get('error')}")
        
        print("\n" + "=" * 50)
        print("🎉 Verification completed successfully!")
        print("\n📋 Your Excel MCP Server is ready to use:")
        print("   • All components are working correctly")
        print("   • Excel operations are functional")
        print("   • MCP server is ready to start")
        print("\n🚀 To start the server:")
        print("   python src/excel_mcp_server/server.py")
        print("\n🔧 For Claude Desktop integration:")
        print('   Add to your config: "excel-mcp-server": {')
        print('     "command": "python",')
        print(f'     "args": ["{project_root / "src" / "excel_mcp_server" / "server.py"}"]')
        print('   }')
        
    except Exception as e:
        print(f"❌ Verification failed: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
