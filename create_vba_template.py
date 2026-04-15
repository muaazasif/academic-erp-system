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
    Application.EnableEvents = True
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Instructions")
    If Not ws Is Nothing Then
        ' Set flag for auto-grader to know macros are enabled
        ws.Range("Z99").Value = "MACROS_OK"
        ws.Range("Z99").Font.Color = RGB(255, 255, 255)
        ws.Range("Z100").Value = ""
        
        ' Unhide all sheets
        Dim s As Worksheet
        For Each s In ThisWorkbook.Worksheets
            If s.Name <> "Instructions" Then s.Visible = xlSheetVisible
        Next s
        ws.Activate
        
        ' Show Motivational Quote
        Dim quotes(0 To 4) As String
        quotes(0) = "🌟 Believe in yourself and all that you are!"
        quotes(1) = "🚀 The only way to do great work is to love what you do."
        quotes(2) = "💪 Success is not final, failure is not fatal: it is the courage to continue that counts."
        quotes(3) = "🔥 Don't watch the clock; do what it does. Keep going!"
        quotes(4) = "✨ Your limitation—it's only your imagination."
        
        Randomize
        MsgBox quotes(Int(5 * Rnd)) & vbCrLf & vbCrLf & "Best of luck for your Excel Exercise!", vbInformation, "Motivational Boost"
    End If
End Sub
"""
        
        # Add code to ThisWorkbook module
        vba_module = workbook.VBProject.VBComponents("ThisWorkbook").CodeModule
        vba_module.AddFromString(vba_code)
        
        # Save as .xlsm (Excel 52 = xlOpenXMLWorkbookMacroEnabled)
        # Use DisplayAlerts = False to overwrite without prompt
        excel.DisplayAlerts = False
        workbook.SaveAs(OUTPUT_PATH, FileFormat=52)
        excel.DisplayAlerts = True
        
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
    # Always recreate the template to update VBA code
    print("🔄 Updating VBA Template...")
    success = create_template_automatically()
    if success:
        print("\n📋 Now run these commands to push to GitHub:")
        print('git add excel_template.xlsm')
        print('git commit -m "Update VBA template with motivational quotes"')
        print('git push origin main')
    else:
        print("\n❌ Failed to update template!")
        print("\n📋 MANUAL STEPS:")
        print("1. Open Excel")
        print("2. ALT+F11")
        print("3. Double-click ThisWorkbook")
        print("4. Paste VBA code")
        print("5. CTRL+S → Save as .xlsm")
        print("6. Save in: E:\\Governor Sindh Course\\Application\\myerp_app\\")
