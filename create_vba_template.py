"""
Create VBA Template AUTOMATICALLY using Windows Excel COM automation
Run this ONCE to create excel_template.xlsm
"""
import os
import sys

OUTPUT_PATH = os.path.join(os.path.dirname(__file__), 'excel_template.xlsm')

def create_template_automatically():
    """Use Windows Excel COM to create .xlsm with VBA macro"""
    try:
        import win32com.client as win32
        
        print("=" * 70)
        print("🔧 Creating Excel Template with VBA...")
        print("=" * 70)
        
        # Open Excel application
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False  # Don't show Excel window
        
        # Create new workbook
        workbook = excel.Workbooks.Add()
        
        # Add VBA code to ThisWorkbook
        vba_code = """Private Sub Workbook_Open()
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
"""
        
        # Add code to ThisWorkbook module
        vba_module = workbook.VBProject.VBComponents("ThisWorkbook").CodeModule
        vba_module.AddFromString(vba_code)
        
        # Save as .xlsm (Excel 52 = xlOpenXMLWorkbookMacroEnabled)
        workbook.SaveAs(OUTPUT_PATH, FileFormat=52)
        
        # Close workbook
        workbook.Close(SaveChanges=False)
        excel.Quit()
        
        print(f"\n✅ SUCCESS! Template created!")
        print(f"📁 Location: {OUTPUT_PATH}")
        print(f"\n🚨 Cheating Detection: ENABLED")
        print("=" * 70)
        
        return True
        
    except ImportError:
        print("❌ Error: win32com not installed!")
        print("\n🔧 Install it with:")
        print("pip install pywin32")
        print("\nThen run this script again!")
        return False
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("\n⚠️ Make sure Microsoft Excel is installed on this PC!")
        return False

if __name__ == '__main__':
    # Check if file already exists
    if os.path.exists(OUTPUT_PATH):
        print("✅ Template already exists!")
        print(f"📁 {OUTPUT_PATH}")
    else:
        success = create_template_automatically()
        if success:
            print("\n📋 Now run these commands to push to GitHub:")
            print('git add excel_template.xlsm')
            print('git commit -m "Add anti-cheating VBA template"')
            print('git push origin main')
        else:
            print("\n❌ Failed to create template!")
            print("\n📋 MANUAL STEPS:")
            print("1. Open Excel")
            print("2. ALT+F11")
            print("3. Double-click ThisWorkbook")
            print("4. Paste VBA code")
            print("5. CTRL+S → Save as .xlsm")
            print("6. Save in: E:\\Governor Sindh Course\\Application\\myerp_app\\")
