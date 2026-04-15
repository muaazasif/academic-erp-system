"""
Setup script to create Excel template with ANTI-CHEATING VBA macro
Run this ONCE to create excel_template.xlsm
"""
import os

VBA_CODE = '''
Private Sub Workbook_Open()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ' Mark that macros are enabled
        ws.Range("Z99").Value = "MACROS_OK"
        ws.Range("Z99").Font.Color = RGB(255, 255, 255) ' White text to hide it
        
        ' Reset cheating flag at start
        ws.Range("Z100").Value = ""
        
        ' Show all other sheets only when macros are enabled
        Dim s As Worksheet
        For Each s In ThisWorkbook.Worksheets
            If s.Name <> "Instructions" Then
                s.Visible = xlSheetVisible
            End If
        Next s
        
        ' Force Instructions to be active
        ws.Activate
    End If
End Sub

Private Sub Workbook_Deactivate()
    DetectCheating
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    DetectCheating
End Sub

Sub DetectCheating()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ' If they haven't been caught yet, catch them now
        If ws.Range("Z100").Value <> "CHEATED" Then
            ws.Range("Z100").Value = "CHEATED"
            ws.Range("Z100").Font.Color = RGB(255, 255, 255)
            
            ' 1. Show Popup Warning
            MsgBox "🚨 CHEATING DETECTED!" & vbCrLf & vbCrLf & _
                   "You switched windows or opened another application." & vbCrLf & _
                   "This attempt has been flagged and your marks will be ZERO.", _
                   vbCritical, "Academic Integrity Warning"
            
            ' 2. Save and Close immediately so they can't change anything
            ThisWorkbook.Save
            ThisWorkbook.Close SaveChanges:=True
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ' Hide all sheets except instructions before saving
    ' This way, if they open it without macros next time, they see nothing
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name <> "Instructions" Then
            s.Visible = xlSheetVeryHidden
        End If
    Next s
    ThisWorkbook.Save
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
