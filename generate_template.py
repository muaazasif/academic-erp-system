"""
Generate Excel Template with Anti-Cheating VBA Macro
Run this to create excel_template.xlsm
"""
import os
import zipfile
import shutil
from xml.etree.ElementTree import Element, SubElement, tostring, ElementTree
import tempfile

def create_xlsm_with_vba(output_path):
    """Create .xlsm file with VBA macro for anti-cheating"""
    
    print("=" * 70)
    print("📝 Creating Excel Template with Anti-Cheating...")
    print("=" * 70)
    
    # VBA Code to detect cheating
    vba_code = '''Attribute VB_Name = "ThisWorkbook"
Private Sub Workbook_Deactivate()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ws.Range("Z100").Value = "CHEATED"
        ws.Range("Z100").Font.Color = RGB(255, 255, 255)
    End If
End Sub
'''
    
    # Try to create a basic .xlsm structure
    try:
        # Create a temporary xlsx file first
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Instructions"
        ws['A1'] = "Excel Assignment"
        
        # Save as xlsx temporarily
        temp_xlsx = output_path.replace('.xlsm', '.xlsx')
        wb.save(temp_xlsx)
        
        # Convert to .xlsm by changing extension and adding VBA
        # Note: .xlsm is actually a .zip file with specific structure
        xlsm_path = output_path
        
        # Rename .xlsx to .xlsm and add VBA project
        # This requires xlwt or xlwings, so we'll provide instructions instead
        
        print("\n⚠️  Python cannot create .xlsm with VBA directly")
        print("\n📋 PLEASE CREATE MANUALLY (takes 1 minute):")
        print("\n1. Open Microsoft Excel")
        print("2. Create new workbook")
        print("3. Press ALT + F11")
        print("4. Double-click 'ThisWorkbook'")
        print("5. Paste this code:")
        print("\n" + "=" * 70)
        print(vba_code)
        print("=" * 70)
        print("\n6. Press CTRL + S")
        print("7. Save as 'excel_template.xlsm'")
        print(f"8. Save in: {output_path}")
        print("\n" + "=" * 70)
        print("✅ Done! System will detect if student opens other files!")
        print("=" * 70)
        
        # Clean up temp xlsx
        if os.path.exists(temp_xlsx):
            os.unlink(temp_xlsx)
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == '__main__':
    output_path = os.path.join(os.path.dirname(__file__), 'excel_template.xlsm')
    create_xlsm_with_vba(output_path)
