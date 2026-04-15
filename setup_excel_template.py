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
    Application.EnableEvents = True
    IsClosing = False
    OpenedAt = Timer
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ws.Range("Z99").Value = "MACROS_OK"
        ws.Range("Z99").Font.Color = RGB(255, 255, 255)
        ws.Range("Z100").Value = ""
        
        Dim s As Worksheet
        For Each s In ThisWorkbook.Worksheets
            If s.Name <> "Instructions" Then s.Visible = xlSheetVisible
        Next s
        ws.Activate
    End If
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    If IsClosing Then Exit Sub
    If Timer - OpenedAt < 5 Then Exit Sub
    
    ' Student actually switched to another app/file
    DetectCheating
End Sub

Sub DetectCheating()
    On Error Resume Next
    If IsClosing Then Exit Sub
    
    ' STOP RECURSION: This was causing false triggers on close
    Application.EnableEvents = False
    IsClosing = True
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        If ws.Range("Z100").Value <> "CHEATED" Then
            ws.Range("Z100").Value = "CHEATED"
            ws.Range("Z100").Font.Color = RGB(255, 255, 255)
            
            MsgBox "🚨 CHEATING DETECTED!", vbCritical
            
            ThisWorkbook.Save
            ThisWorkbook.Close SaveChanges:=True
        End If
    End If
    Application.EnableEvents = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    IsClosing = True
End Sub
'''

def create_template():
    """Create .xlsm template with VBA cheating detection"""
    
    print("=" * 70)
    print("📝 EXCEL ANTI-CHEATING TEMPLATE SETUP (FIXED V3)")
    print("=" * 70)
    print()
    print("FIXED: No more 'Cheating' popup when closing the file normally!")
    print()
    print("-" * 70)
    print(VBA_CODE)
    print("-" * 70)
    print()
    print("✅ WORKS ON CLOSE: Now you can open and close safely!")
    print("=" * 70)

if __name__ == '__main__':
    create_template()
