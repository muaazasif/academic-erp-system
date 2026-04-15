@echo off
echo ============================================================
echo    EXCEL ANTI-CHEATING TEMPLATE SETUP
echo ============================================================
echo.
echo 1. Opening Excel...
echo.
timeout /t 2
start excel
echo.
echo 2. Excel is opening!
echo.
echo NOW DO THIS:
echo.
echo    a) Click "Blank workbook"
echo    b) Press ALT + F11 (VBA Editor opens)
echo    c) Double-click "ThisWorkbook" (left panel)
echo    d) Paste this code:
echo.
echo    --------------------------------------------------------
echo    Private Sub Workbook_Deactivate()
echo        On Error Resume Next
echo        Dim ws As Worksheet
echo        Set ws = ThisWorkbook.Sheets("Instructions")
echo        If Not ws Is Nothing Then
echo            ws.Range("Z100").Value = "CHEATED"
echo            ws.Range("Z100").Font.Color = RGB(255, 255, 255)
echo        End If
echo    End Sub
echo    --------------------------------------------------------
echo.
echo    e) Press CTRL + S
echo    f) Save as "Excel Macro-Enabled Workbook (*.xlsm)"
echo    g) Save in: %~dp0
echo    h) File name: excel_template
echo.
echo 3. After saving, run this to commit to GitHub:
echo    git add excel_template.xlsm
echo    git commit -m "Add anti-cheating VBA template"
echo    git push origin main
echo.
echo ============================================================
pause
