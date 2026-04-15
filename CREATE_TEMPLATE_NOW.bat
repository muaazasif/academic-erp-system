@echo off
echo ============================================================
echo    EXCEL ANTI-CHEATING TEMPLATE SETUP (FIXED V2)
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
echo    b) Create a sheet named "Instructions" (exactly)
echo    c) Press ALT + F11 (VBA Editor opens)
echo    d) Double-click "ThisWorkbook" (left panel)
echo    e) Paste this code (REPLACE EVERYTHING):
echo.
echo    --------------------------------------------------------
echo    Dim IsClosing As Boolean
echo    Dim OpenedAt As Double
echo
echo    Private Sub Workbook_Open()
echo        On Error Resume Next
echo        IsClosing = False
echo        OpenedAt = Timer
echo        Dim ws As Worksheet
echo        Set ws = ThisWorkbook.Sheets("Instructions")
echo        If Not ws Is Nothing Then
echo            ws.Range("Z99").Value = "MACROS_OK"
echo            ws.Range("Z100").Value = ""
echo            Dim s As Worksheet
echo            For Each s In ThisWorkbook.Worksheets
echo                If s.Name ^<^> "Instructions" Then s.Visible = xlSheetVisible
echo            Next s
echo            ws.Activate
echo        End If
echo    End Sub
echo
echo    Private Sub Workbook_Deactivate()
echo        On Error Resume Next
echo        ' 5 SECOND GRACE PERIOD TO ALLOW MACRO ENABLING
echo        If Timer - OpenedAt ^< 5 Then Exit Sub
echo        If Not IsClosing Then DetectCheating
echo    End Sub
echo
echo    Sub DetectCheating()
echo        On Error Resume Next
echo        If IsClosing Then Exit Sub
echo        Dim ws As Worksheet
echo        Set ws = ThisWorkbook.Sheets("Instructions")
echo        If Not ws Is Nothing Then
echo            If ws.Range("Z100").Value ^<^> "CHEATED" Then
echo                ws.Range("Z100").Value = "CHEATED"
echo                MsgBox "🚨 CHEATING DETECTED!", vbCritical
echo                IsClosing = True
echo                ThisWorkbook.Save
echo                ThisWorkbook.Close SaveChanges:=True
echo            End If
echo        End If
echo    End Sub
echo
echo    Private Sub Workbook_BeforeClose(Cancel As Boolean)
echo        IsClosing = True
echo    End Sub
echo    --------------------------------------------------------
echo.
echo    f) Press CTRL + S
echo    g) Save as "Excel Macro-Enabled Workbook (*.xlsm)"
echo    h) File name: excel_template
echo.
echo ✅ FIXED: 5-second grace period added to prevent false positives!
echo.
pause
