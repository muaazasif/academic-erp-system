"""
Setup script to create Excel template with ANTI-CHEATING VBA macro
Run this ONCE to create excel_template.xlsm
"""
import os

VBA_CODE = '''
Private Sub Workbook_Deactivate()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ws.Range("Z100").Value = "CHEATED"
        ws.Range("Z100").Font.Color = RGB(255, 255, 255)
    End If
End Sub

Private Sub Workbook_Activate()
    ' Do nothing when coming back
End Sub
'''

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'excel_template.xlsm')

def create_template():
    """Create .xlsm template with VBA cheating detection"""
    
    print("=" * 70)
    print("📝 EXCEL ANTI-CHEATING TEMPLATE SETUP")
    print("=" * 70)
    print()
    print("Follow these steps to create the template:")
    print()
    print("1️⃣ Open Microsoft Excel")
    print("2️⃣ Create a NEW blank workbook")
    print("3️⃣ Press ALT + F11 (opens VBA Editor)")
    print("4️⃣ In Project Explorer (left side), double-click 'ThisWorkbook'")
    print("5️⃣ Paste this EXACT code:")
    print()
    print("-" * 70)
    print(VBA_CODE)
    print("-" * 70)
    print()
    print("6️⃣ Press CTRL + S (Save)")
    print("7️⃣ Choose 'Excel Macro-Enabled Workbook (*.xlsm)'")
    print("8️⃣ Save as: excel_template.xlsm")
    print("9️⃣ Save in THIS folder:")
    print(f"   {os.path.dirname(__file__)}")
    print()
    print("=" * 70)
    print("✅ Once done, the system will detect cheating!")
    print("🚨 If student opens ANY other file/window → MARKS = ZERO")
    print("=" * 70)
    
    # Check if template exists
    if os.path.exists(TEMPLATE_PATH):
        print("\n✅ Template file found!")
    else:
        print(f"\n❌ Template NOT found at: {TEMPLATE_PATH}")
        print("📋 Please create it using steps above!")

if __name__ == '__main__':
    create_template()
