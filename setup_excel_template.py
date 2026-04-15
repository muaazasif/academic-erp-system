"""
Setup script to create Excel template with ANTI-CHEATING VBA macro
Run this ONCE to create excel_template.xlsm
"""
import os

VBA_CODE = '''
Dim IsClosing As Boolean
Dim OpenedAt As Double

Private Sub Workbook_Open()
    On Error Resume Next
    IsClosing = False
    OpenedAt = Timer ' Record time of opening
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ' Mark that macros are enabled
        ws.Range("Z99").Value = "MACROS_OK"
        ws.Range("Z99").Font.Color = RGB(255, 255, 255)
        
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
    On Error Resume Next
    ' GRACE PERIOD: Ignore deactivate for first 5 seconds
    ' This prevents false cheating trigger when clicking "Enable Content"
    If Timer - OpenedAt < 5 Then Exit Sub
    
    If Not IsClosing Then DetectCheating
End Sub

Sub DetectCheating()
    On Error Resume Next
    If IsClosing Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        If ws.Range("Z100").Value <> "CHEATED" Then
            ws.Range("Z100").Value = "CHEATED"
            ws.Range("Z100").Font.Color = RGB(255, 255, 255)
            
            MsgBox "🚨 CHEATING DETECTED!" & vbCrLf & vbCrLf & _
                   "You switched to another window or application." & vbCrLf & _
                   "This attempt is flagged and marks will be ZERO.", _
                   vbCritical, "Academic Integrity Warning"
            
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

def create_template():
    """Create .xlsm template with VBA cheating detection"""
    
    print("=" * 70)
    print("📝 EXCEL ANTI-CHEATING TEMPLATE SETUP (FIXED)")
    print("=" * 70)
    print()
    print("This version fixes the 'Automatic Cheating' bug on open.")
    print()
    print("-" * 70)
    print(VBA_CODE)
    print("-" * 70)
    print()
    print("✅ GRACE PERIOD ADDED: Now you have 5 seconds to click 'Enable Macros' safely!")
    print("=" * 70)

if __name__ == '__main__':
    create_template()
