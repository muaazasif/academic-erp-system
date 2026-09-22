import openpyxl

def create_new_template():
    wb = openpyxl.Workbook()
    
    # Sheet 1: Instructions (Required for Anti-Cheating)
    ws_instr = wb.active
    ws_instr.title = "Instructions"
    ws_instr['A1'] = "EXCEL SKILL 5 ASSIGNMENT"
    ws_instr['A2'] = "Ensure Macros are enabled to avoid automatic closure."
    ws_instr['Z99'] = "MACROS_OK" # Mocked trigger for VBA
    
    # Sheet 2: Assignment
    ws_assign = wb.create_sheet(title="Assignment")
    ws_assign['A10'] = "VLOOKUP Result"
    ws_assign['A11'] = "SUMIF Result"
    ws_assign['A12'] = "COUNTIF Result"
    ws_assign['A13'] = "IF Formula Result"
    
    # Placeholders
    ws_assign['B10'] = "..."
    ws_assign['B11'] = "..."
    ws_assign['B12'] = "..."
    ws_assign['B13'] = "..."
    
    # Save as .xlsx (VBA requires .xlsm, but we start with structure)
    # The user can then add the VBA macro or we can provide a method to merge
    wb.save('static/solutions/Excel_Skill_5_Template.xlsx')
    print("✅ New Excel template created at static/solutions/Excel_Skill_5_Template.xlsx")

if __name__ == '__main__':
    create_new_template()
