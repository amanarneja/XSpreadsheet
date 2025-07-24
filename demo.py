#!/usr/bin/env python3
"""
Excel MCP Server Demo

This script demonstrates how to use the Excel MCP Server functionality.
It creates sample Excel files and shows various operations.
"""

import sys
import os
from pathlib import Path

# Add the project to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root / "src" / "excel_mcp_server"))

from excel_library import ExcelLibrary

def create_sample_data():
    """Create sample data for demonstration"""
    return [
        ["Product", "Category", "Price", "Stock", "Sales"],
        ["Laptop", "Electronics", 999.99, 50, 120],
        ["Mouse", "Electronics", 29.99, 200, 350],
        ["Keyboard", "Electronics", 79.99, 150, 280],
        ["Desk", "Furniture", 299.99, 25, 45],
        ["Chair", "Furniture", 199.99, 30, 67],
        ["Book", "Education", 19.99, 100, 89],
        ["Pen", "Stationery", 2.99, 500, 756],
        ["Notebook", "Stationery", 5.99, 300, 423]
    ]

def demo_basic_operations():
    """Demonstrate basic Excel operations"""
    print("📊 Excel MCP Server Demo")
    print("=" * 50)
    
    # Initialize the Excel library
    excel_lib = ExcelLibrary()
    
    # Create sample data
    sample_data = create_sample_data()
    headers = sample_data[0]
    data_rows = sample_data[1:]
    
    print("1. Creating Excel file with sample data...")
    
    # Write data to Excel file
    result = excel_lib.write_excel_file(
        "demo_inventory.xlsx",
        data_rows,
        sheet_name="Inventory",
        headers=headers
    )
    if result.get('success'):
        print(f"   ✅ Created Excel file with {result['rows_written']} rows")
    else:
        print(f"   ❌ Error: {result.get('error', 'Unknown error')}")
    
    print("\n2. Reading data back from Excel file...")
    
    # Read the data back
    read_result = excel_lib.read_excel_file("demo_inventory.xlsx", "Inventory")
    print(f"   ✅ Read {len(read_result['data'])} rows of data")
    
    print("\n3. Getting worksheet information...")
    
    # Get worksheet info
    info_result = excel_lib.get_worksheet_info("demo_inventory.xlsx")
    print(f"   ✅ Workbook contains {len(info_result['worksheets'])} worksheet(s)")
    for ws in info_result['worksheets']:
        print(f"      - {ws['name']}: {ws['max_row']} rows x {ws['max_column']} columns")
    
    print("\n4. Adding a new worksheet...")
    
    # Add a new worksheet
    add_result = excel_lib.add_worksheet("demo_inventory.xlsx", "Summary")
    if add_result.get('success'):
        print(f"   ✅ Added worksheet: {add_result.get('sheet_name', 'Summary')}")
    else:
        print(f"   ❌ Error: {add_result.get('error', 'Unknown error')}")
    
    print("\n5. Updating a specific cell...")
    
    # Update a cell
    update_result = excel_lib.update_cell(
        "demo_inventory.xlsx", 
        "Summary", 
        "A1", 
        "Inventory Summary Report"
    )
    if update_result.get('success'):
        print(f"   ✅ Updated cell A1")
    else:
        print(f"   ❌ Error: {update_result.get('error', 'Unknown error')}")
    
    # Add a label for the formula
    label_result = excel_lib.update_cell(
        "demo_inventory.xlsx", 
        "Summary", 
        "A2", 
        "Total Stock:"
    )
    
    print("\n6. Applying a formula...")
    
    # Apply a formula
    formula_result = excel_lib.apply_formula(
        "demo_inventory.xlsx",
        "Summary",
        "A3",
        "=SUM(Inventory!D:D)"
    )
    if formula_result.get('success'):
        print(f"   ✅ Applied formula to cell A3")
    else:
        print(f"   ❌ Error: {formula_result.get('error', 'Unknown error')}")
    
    print("\n8. Applying cell formatting...")
    
    # Format cells
    format_result = excel_lib.format_cells(
        "demo_inventory.xlsx",
        "Summary",
        "A1:A3",
        {
            "bold": True,
            "font_size": 14,
            "background_color": "FFADD8E6"
        }
    )
    if format_result.get('success'):
        print(f"   ✅ Applied formatting to range A1:A3")
    else:
        print(f"   ❌ Error: {format_result.get('error', 'Unknown error')}")
    
    print(f"\n🎉 Demo completed! Check 'demo_inventory.xlsx' for results.")
    print("\nThe file contains:")
    print("  - 'Inventory' sheet with sample product data")
    print("  - 'Summary' sheet with formatted title and total stock formula")
    
    return "demo_inventory.xlsx"

def demo_chart_creation():
    """Demonstrate chart creation"""
    print("\n📈 Chart Creation Demo")
    print("=" * 30)
    
    excel_lib = ExcelLibrary()
    
    # Create sales data for chart
    sales_data = [
        ["Month", "Sales"],
        ["Jan", 15000],
        ["Feb", 18000],
        ["Mar", 22000],
        ["Apr", 19000],
        ["May", 25000],
        ["Jun", 28000]
    ]
    
    print("1. Creating sales data...")
    
    # Write sales data
    excel_lib.write_excel_file(
        "demo_charts.xlsx",
        sales_data[1:],
        sheet_name="SalesData",
        headers=sales_data[0]
    )
    print("   ✅ Sales data created")
    
    print("2. Creating line chart...")
    
    # Create a line chart
    chart_result = excel_lib.create_chart(
        "demo_charts.xlsx",
        "SalesData",
        "A1:B7",
        "line",
        "Monthly Sales Trend",
        "D2"
    )
    if chart_result.get('success'):
        print(f"   ✅ Created line chart")
    else:
        print(f"   ❌ Error: {chart_result.get('error', 'Unknown error')}")
    
    print(f"\n🎉 Chart demo completed! Check 'demo_charts.xlsx' for the sales chart.")
    
    return "demo_charts.xlsx"

def cleanup_demo_files():
    """Clean up demo files"""
    demo_files = ["demo_inventory.xlsx", "demo_charts.xlsx"]
    
    print(f"\n🧹 Cleanup")
    print("=" * 15)
    
    for file in demo_files:
        if os.path.exists(file):
            os.remove(file)
            print(f"   ✅ Removed {file}")
        else:
            print(f"   ℹ️  {file} not found")

def main():
    """Main demo function"""
    try:
        # Run basic operations demo
        demo_basic_operations()
        
        # Run chart creation demo
        demo_chart_creation()
        
        print(f"\n" + "=" * 50)
        print("✨ All demos completed successfully!")
        print("\nYour Excel MCP Server is working perfectly and ready to use.")
        print("\nTo start the MCP server:")
        print("  python src/excel_mcp_server/server.py")
        print("\nTo use with Claude Desktop, add this to your config:")
        print('  "excel-mcp-server": {')
        print('    "command": "python",')
        print(f'    "args": ["{project_root / "src" / "excel_mcp_server" / "server.py"}"]')
        print('  }')
        
        # Ask if user wants to keep demo files
        response = input(f"\nKeep demo files for inspection? (y/n): ").lower().strip()
        if response not in ['y', 'yes']:
            cleanup_demo_files()
        else:
            print("📁 Demo files kept for your inspection.")
            
    except Exception as e:
        print(f"\n❌ Demo failed: {e}")
        print("Please check that all dependencies are installed correctly.")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
