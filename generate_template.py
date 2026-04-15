"""
Generate Excel Template with Anti-Cheating VBA Macro
"""
import os
import shutil
import openpyxl

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
        
        ws.Activate
    End If
End Sub

Private Sub Workbook_Deactivate()
    If Not IsClosing Then DetectCheating
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
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
            MsgBox "🚨 CHEATING DETECTED!", vbCritical
            ThisWorkbook.Save
            ThisWorkbook.Close SaveChanges:=True
        End If
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    IsClosing = True
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name <> "Instructions" Then
            s.Visible = xlSheetVeryHidden
        End If
    Next s
    ThisWorkbook.Save
End Sub
'''

def setup_template():
    output_path = 'excel_template.xlsm'
    print(f"📝 Creating Template Instructions for: {output_path}")
    print("\nMANUAL STEPS REQUIRED:")
    print("1. Open Excel")
    print("2. Create Sheets: Instructions, VLOOKUP, SUMIF & COUNTIF, LEFT RIGHT MID, IF & NESTED IF, COMPLEX CHALLENGE")
    print("3. Press ALT+F11")
    print("4. Double-click 'ThisWorkbook'")
    print("5. Paste this code:")
    print("-" * 50)
    print(VBA_CODE)
    print("-" * 50)
    print("6. Save as 'excel_template.xlsm'")

if __name__ == "__main__":
    setup_template()
