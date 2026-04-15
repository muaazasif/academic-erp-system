"""
Setup script to create Excel template with ANTI-CHEATING VBA macro
Run this ONCE to create excel_template.xlsm
"""
import os

VBA_CODE = '''
Dim IsClosing As Boolean

Private Sub Workbook_Open()
    On Error Resume Next
    IsClosing = False
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

' THIS IS THE ONLY EVENT WE NEED
Private Sub Workbook_Deactivate()
    On Error Resume Next
    If Not IsClosing Then 
        DetectCheating
    End If
End Sub

Sub DetectCheating()
    On Error Resume Next
    If IsClosing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ' If they haven't been caught yet, catch them now
        If ws.Range("Z100").Value <> "CHEATED" Then
            ws.Range("Z100").Value = "CHEATED"
            ws.Range("Z100").Font.Color = RGB(255, 255, 255)
            
            ' 1. Show Popup Warning
            MsgBox "🚨 CHEATING DETECTED!" & vbCrLf & vbCrLf & _
                   "You switched to another window or application." & vbCrLf & _
                   "Your attempt is flagged and marks will be ZERO.", _
                   vbCritical, "Academic Integrity Warning"
            
            ' 2. Save and Close immediately
            IsClosing = True
            ThisWorkbook.Save
            ThisWorkbook.Close SaveChanges:=True
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    IsClosing = True
    ' Hide all sheets except instructions before saving
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
    print("📝 EXCEL ANTI-CHEATING TEMPLATE SETUP (FIXED)")
    print("=" * 70)
    print()
    print("Follow these steps exactly:")
    print()
    print("1️⃣ Open Microsoft Excel -> New Blank Workbook")
    print("2️⃣ Create a sheet named 'Instructions' (Case Sensitive)")
    print("3️⃣ Press ALT + F11 (VBA Editor)")
    print("4️⃣ Double-click 'ThisWorkbook' on the left side")
    print("5️⃣ Paste this EXACT code (Overwrite everything):")
    print()
    print("-" * 70)
    print(VBA_CODE)
    print("-" * 70)
    print()
    print("6️⃣ Press CTRL + S -> Save as 'Excel Macro-Enabled Workbook (*.xlsm)'")
    print("7️⃣ File name: excel_template.xlsm")
    print()
    print("✅ NOW SHEET SWITCHING WILL NOT TRIGGER CHEATING!")
    print("🚨 ONLY SWITCHING TO CHROME/OTHER FILES WILL TRIGGER ZERO MARKS.")
    print("=" * 70)

if __name__ == '__main__':
    create_template()
