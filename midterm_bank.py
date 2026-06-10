import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import random
import re

# ============================================
# MASTER TASK BANK (100 TASKS) - EXPERT MODE
# ============================================

def get_unique_style(tid):
    """Returns a unique color scheme and offset based on TID"""
    colors = [
        "1F4E79", "C00000", "2E75B6", "7030A0", "375623", 
        "BF8F00", "0070C0", "843C0C", "548235", "3B3838"
    ]
    # Rotate colors and offsets to ensure visual variety
    return {
        'color': colors[tid % len(colors)],
        'row_offset': (tid % 5),
        'col_offset': (tid % 3)
    }

def style_header(ws, row, col_start, cols, color="1F4E79"):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    font = Font(bold=True, color="FFFFFF")
    for c in range(col_start, col_start + cols):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

def get_task_bank():
    """Returns 100 tasks with HIGH visual and logical variety"""
    bank = {}
    
    problems = {
        1: "Regional performance: If Sales > 50000 and Profit > 20%, status 'ELITE'. If Sales > 30000, 'PRIME'. Else 'standard'.",
        2: "Extract Domain: Use MID and FIND to isolate domain name from emails like 'user@sub.company.org'.",
        3: "Working Days: Calculate completion date 120 working days after today, excluding holidays in G4:G10.",
        4: "Nested IF: Grade students: >95:A+, >90:A, >85:B+, >80:B, >70:C, Else:F.",
        5: "Text Cleaning: Use TRIM, PROPER, and SUBSTITUTE to clean '  mUAAZ_asif-PAKISTAN  ' into 'Muaaz Asif Pakistan'.",
        6: "Array Sum: Total sales for 'Electronics' ONLY for 'North' region after '2024-03-01'.",
        7: "Logic Gate: If (A AND B) OR (C AND NOT D), then 1, else 0. 50 rows of data.",
        8: "Data Validation: Column B only accepts even numbers between 500 and 1500.",
        9: "Dynamic Offset: Use OFFSET to sum the last 5 months of sales data automatically.",
        10: "Error Handling: Use IFERROR and VLOOKUP; if not found, show 'ID_MISSING_CONTACT_ADMIN'.",
        11: "Conditional Formatting: Highlight rows where Profit is negative AND Sales are in the top 10%.",
        12: "Loan Amortization: Total interest paid over 15 years for loan of 2M at 7.5% fixed.",
        13: "Date Diff: Calculate age in Years, Months, and Days exactly from birthdate in cell C5.",
        14: "Unique Count: Array formula to count unique student names in list of 500 entries.",
        15: "Ranking: Rank products on Sales; if tied, use Profit % as tie-breaker.",
        16: "Data Shifting: Move all non-empty values from a column to new range, removing blanks.",
        17: "Multi-Sheet Lookup: VLOOKUP value across 3 sheets (SheetA, SheetB, SheetC) until found.",
        18: "Frequency: Distribution of test scores into buckets of 10 (0-10, 11-20...).",
        19: "Text Joining: Combine Name, City, Zip into 'NAME [CITY] - ZIP' for 100 users.",
        20: "Transposing: Take 5x5 matrix and rotate it 90 degrees using formulas only.",
        21: "Pro Lookup: INDEX/MATCH/MATCH for 2-way matrix lookup (Product vs Month).",
        22: "Dynamic Filter: Filter transactions to show only those > $10,000 for 'Refund' category.",
        23: "Sorting: Sort dynamic range by Category (ASC) and then by Date (DESC) using formulas.",
        24: "XLOOKUP Power: XLOOKUP to find first occurrence from bottom of 1000-row list.",
        25: "Scenario Manager: Model that switches between 'Best Case', 'Worst Case', and 'Realistic'.",
        26: "Goal Seek: Required sales volume to reach target profit of $50,000.",
        27: "Solver: Optimize mix of 3 products for max profit given 40-hour labor constraint.",
        28: "Indirect Ref: Summary sheet pulling 'Total' from sheets named by Month (Jan, Feb...).",
        29: "Recursive Lambda: Recursive logic to calculate Fibonacci sequence up to 20 terms.",
        30: "Data Table: 2-variable data table for Interest Rate vs. Term in mortgage model.",
        31: "Match Type: MATCH with '-1' for smallest value greater than or equal to target.",
        32: "Wildcard VLOOKUP: Partial product names like '*Pro*' in database of 500 items.",
        33: "Dynamic Range: Named Range that automatically expands as new rows are added.",
        34: "Subtotals: SUBTOTAL(9, ...) to sum filtered data only, ignoring hidden rows.",
        35: "Array Constants: Use {1,2,3;4,5,6} for matrix multiplication in a single cell.",
        36: "Fiscal Year: Convert dates into Fiscal Year quarters (starting July 1st).",
        37: "Network Days: Working hours between two dates, accounting for 8-hour workdays.",
        38: "Text to Columns: Split CSV string into 5 separate cells using formulas only.",
        39: "Duplicate Finder: Identify IDs that appear more than 3 times in a list.",
        40: "Z-Score: Z-Score for set of student marks to identify outliers.",
        41: "Joins: INNER JOIN Students, Courses, and Instructors to list 'Student Name - Course - Instructor'.",
        42: "Aggregation: Top 3 Departments by total revenue in last 6 months.",
        43: "Subqueries: Select employees with salary higher than average of their department.",
        44: "Window Functions: Rank sales reps within each region using ROW_NUMBER() and PARTITION BY.",
        45: "CTEs: WITH clause to calculate recursive category hierarchy.",
        46: "Case Statements: Categorize transactions as 'Small' (<100), 'Medium' (100-500), 'Large' (>500).",
        47: "Exists: Find customers who ordered in 2023 but NO orders in 2024.",
        48: "Grouping Sets: GROUPING SETS for subtotal of (Region, Category) and (Region).",
        49: "Strings: COALESCE and CONCAT for full addresses, handling null 'Unit Number'.",
        50: "Dates: Extract day-of-week and avg sales for every 'Friday' in history.",
        51: "Performance: Query using INDEX hint to optimize search on 'last_name'.",
        52: "Self Join: Employees and Manager's name by joining table to itself.",
        53: "Distinct: Unique cities each student has lived in according to history table.",
        54: "Having: Departments with > 10 employees AND average salary < 40k.",
        55: "Updates: UPDATE query with JOIN to increase prices by 10% for 'Discontinued' items.",
        56: "Deletes: Remove duplicate records from staging, keeping only newest ID.",
        57: "Views: SECURE VIEW hiding 'Social Security Number' from regular users.",
        58: "Math: Standard deviation of 'Wait Time' for customer support tickets.",
        59: "Union: Merge 'Past_Students' and 'Current_Students' while removing duplicates.",
        60: "Cast: Convert 'Text' timestamp into valid 'Date' and calculate months since creation.",
        61: "Pivoting: Convert 12-month column layout into flat 'Attribute/Value' format.",
        62: "Merging: Join local Excel table with web-based JSON API source.",
        63: "Cleaning: Remove non-printable characters and extra spaces from 5000-row column.",
        64: "M-Code: Custom M-function to calculate Tax based on dynamic parameter.",
        65: "Fuzzy Match: Merge customer lists where names are slightly different ('J. Doe' vs 'John Doe').",
        66: "Splitting: Split 'Path' string (C:\\Docs\\File.txt) into Folder and Filename.",
        67: "Parameters: PQ parameter to change 'Source Folder' dynamically.",
        68: "Appending: Combine 12 CSV files from folder into master table automatically.",
        69: "Grouping: Group by Year and Month to find 'Maximum' sale of each month.",
        70: "Conditionals: Custom Column with 5 logic rules using M-code IF.",
        71: "Transposing: Flip headers-on-left table into standard headers-on-top.",
        72: "Replacing: Use lookup table to replace 50 abbreviations with full names.",
        73: "Sorting: Ensure PQ output is sorted by 'Priority' (High, Med, Low) correctly.",
        74: "Fill Down: 'Fill Down' to populate null category cells from row above.",
        75: "Extracting: Isolate text between delimiters (e.g. between '(' and ')').",
        76: "Data Types: Fix 'Scientific Notation' numbers imported as text.",
        77: "Buffers: Table.Buffer to optimize slow merge on large datasets.",
        78: "Distinct: Keep only first occurrence of each 'Customer ID' while sorting by 'Date'.",
        79: "Filtering: Filter rows where 'Comments' contains word from forbidden list.",
        80: "Schema: Force schema on dynamic folder import to prevent 'Column Not Found'.",
        81: "Loops: Loop through 100 sheets and protect them with random password.",
        82: "UserForms: Pop-up form validating 'Email' and 'Phone' before adding to sheet.",
        83: "Events: Workbook_Open macro logging Username and Time to hidden sheet.",
        84: "FileSystem: List all filenames from 'C:\\Midterm_Data' into Column A.",
        85: "Optimization: Macro running 10x faster by disabling ScreenUpdating.",
        86: "Object-Oriented: Class Module 'Student' with properties 'Name' and 'Marks'.",
        87: "Charts: Automate 'Pareto Chart' creation based on selection.",
        88: "Email: Sub to send personalized Outlook email to everyone in list.",
        89: "Outlook: Save all attachments from 'Submissions' folder to local path.",
        90: "RegEx: 'VBScript.RegExp' to validate password (1 cap, 1 num).",
        91: "API: VBA function to fetch USD/PKR exchange rate from public API.",
        92: "Arrays: Read 10,000 rows into Array, process, and write back in one shot.",
        93: "Dictionary: 'Scripting.Dictionary' for unique values and frequencies.",
        94: "Errors: Global error handler logging line number of crash.",
        95: "Formatting: Macro converting selected table into 'Professional PDF'.",
        96: "Protection: Macro hiding 'Formulas Bar' and 'Ribbon' for kiosk mode.",
        97: "Sorting: 'Bubble Sort' algorithm in VBA to sort 1D array.",
        98: "SQL in VBA: ADODB to run SQL against external Access database.",
        99: "VBA Logic: Function calculating 'Shortest Path' in a matrix.",
        100: "Master Macro: Button running Cleaning -> Formatting -> Emailing."
    }

    for tid, instr in problems.items():
        if tid <= 20: gen = make_excel_logic_gen(tid, instr)
        elif tid <= 40: gen = make_excel_advanced_gen(tid, instr)
        elif tid <= 60: gen = make_sql_gen(tid, instr)
        elif tid <= 80: gen = make_power_query_gen(tid, instr)
        else: gen = make_vba_gen(tid, instr)
        
        bank[tid] = {
            'id': tid,
            'topic': get_topic_name(tid),
            'title': f"{get_topic_name(tid).replace('_',' ')} Task {tid}",
            'generate': gen,
            'grade': make_basic_grader(tid)
        }
    return bank

def make_basic_grader(tid):
    """Creates a grader that checks if the yellow cells are filled for a specific task"""
    def grade(wb, sheet_name):
        try:
            if sheet_name not in wb.sheetnames:
                return 0.0
            ws = wb[sheet_name]
            
            # Find the yellow cells (we know where they are based on the generators)
            style = get_unique_style(tid)
            r_start = 2 + style['row_offset']
            c_start = 1 + style['col_offset']
            
            filled_count = 0
            total_expected = 0
            
            if tid <= 20: # EXCEL_LOGIC
                total_expected = 7
                for i in range(1, 8):
                    val = ws.cell(row=r_start+4+i, column=c_start+2).value
                    if val is not None and str(val).strip() != "":
                        filled_count += 1
            elif tid <= 40: # EXCEL_ADVANCED
                total_expected = 2
                if ws.cell(row=r_start+5, column=c_start+1).value: filled_count += 1
                if ws.cell(row=r_start+6, column=c_start+1).value: filled_count += 1
            elif tid <= 60: # SQL_QUERIES
                # SQL box is yellow
                total_expected = 1
                val = ws.cell(row=r_start+5, column=c_start).value
                if val and "-- Write SQL Here --" not in str(val) and str(val).strip() != "":
                    filled_count = 1
            elif tid <= 80: # POWER_QUERY
                # CLEAN DESTINATION column is yellow
                rows = 5 + (tid % 8)
                total_expected = rows - 1
                cols = 3 + (tid % 3)
                for r in range(1, rows):
                    if ws.cell(row=r_start+5+r, column=c_start+cols+1).value:
                        filled_count += 1
            else: # VBA_EXPERT
                # VBA box is yellow
                total_expected = 1
                val = ws.cell(row=r_start+5, column=c_start).value
                if val and "' Sub Task_" not in str(val) and str(val).strip() != "":
                    filled_count = 1
            
            if total_expected == 0: return 1.0 # Should not happen
            return round(filled_count / total_expected, 2)
        except Exception as e:
            print(f"Error grading task {tid} in sheet {sheet_name}: {e}")
            return 0.0
    return grade

def get_topic_name(i):
    if i <= 20: return "EXCEL_LOGIC"
    if i <= 40: return "EXCEL_ADVANCED"
    if i <= 60: return "SQL_QUERIES"
    if i <= 80: return "POWER_QUERY"
    return "VBA_EXPERT"

# ---------------------------------------------------------
# UNIQUE LAYOUT GENERATORS (HIGH VARIETY)
# ---------------------------------------------------------

def make_excel_logic_gen(tid, instruction):
    def generate(ws):
        style = get_unique_style(tid)
        r_start = 2 + style['row_offset']
        c_start = 1 + style['col_offset']
        ws.title = f"Logic_{tid}"
        ws.cell(row=r_start, column=c_start, value=f"🏆 TASK {tid}: EXPERT LOGIC").font = Font(size=14, bold=True, color=style['color'])
        ws.cell(row=r_start+2, column=c_start, value="GOAL:").font = Font(bold=True)
        ws.cell(row=r_start+2, column=c_start+1, value=instruction).font = Font(bold=True)
        
        headers = ['Input_X', 'Input_Y', 'Logic_Test_Result']
        style_header(ws, r_start+4, c_start, 3, color=style['color'])
        for i in range(1, 8):
            ws.cell(row=r_start+4+i, column=c_start, value=random.randint(10, 500))
            ws.cell(row=r_start+4+i, column=c_start+1, value=random.choice(['True','False','Pending']))
            ws.cell(row=r_start+4+i, column=c_start+2).fill = PatternFill(start_color="FFFF00", fill_type="solid")
    return generate

def make_excel_advanced_gen(tid, instruction):
    def generate(ws):
        style = get_unique_style(tid)
        r_start = 3 + style['row_offset']
        c_start = 1 + style['col_offset']
        ws.title = f"Adv_{tid}"
        ws.cell(row=r_start, column=c_start, value=f"🏆 TASK {tid}: DATA MODELING").font = Font(size=14, bold=True, color=style['color'])
        ws.cell(row=r_start+1, column=c_start, value=instruction)
        
        # Matrix Layout
        ws.cell(row=r_start+3, column=c_start+4, value="REFERENCE MATRIX").font = Font(bold=True)
        style_header(ws, r_start+4, c_start+4, 3, color=style['color'])
        for r in range(1, 5):
            for c in range(1, 4):
                ws.cell(row=r_start+4+r, column=c_start+3+c, value=random.randint(1,1000))
        
        ws.cell(row=r_start+4, column=c_start, value="INPUT AREA").font = Font(bold=True)
        ws.cell(row=r_start+5, column=c_start, value="Parameter")
        ws.cell(row=r_start+5, column=c_start+1).fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws.cell(row=r_start+6, column=c_start, value="Solution")
        ws.cell(row=r_start+6, column=c_start+1).fill = PatternFill(start_color="FFFF00", fill_type="solid")
    return generate

def make_sql_gen(tid, instruction):
    def generate(ws):
        style = get_unique_style(tid)
        r_start = 2 + style['row_offset']
        c_start = 1 + style['col_offset']
        ws.title = f"SQL_{tid}"
        ws.cell(row=r_start, column=c_start, value=f"🏆 TASK {tid}: SQL ENGINEERING").font = Font(size=16, bold=True, color=style['color'])
        ws.cell(row=r_start+2, column=c_start, value="QUERY REQ:").font = Font(bold=True)
        ws.cell(row=r_start+2, column=c_start+1, value=instruction)
        ws.cell(row=r_start+4, column=c_start, value="[ CONSOLE BOX ]").font = Font(bold=True)
        # Unique Console Box size for every SQL task
        box_height = 5 + (tid % 5)
        for r in range(r_start+5, r_start+5+box_height):
            for c in range(c_start, c_start+8):
                ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws.cell(row=r_start+5, column=c_start, value="-- Write SQL Here --")
    return generate

def make_power_query_gen(tid, instruction):
    def generate(ws):
        style = get_unique_style(tid)
        r_start = 2 + style['row_offset']
        c_start = 1 + style['col_offset']
        ws.title = f"ETL_{tid}"
        ws.cell(row=r_start, column=c_start, value=f"🏆 TASK {tid}: ETL CLEANING").font = Font(size=14, bold=True, color=style['color'])
        ws.cell(row=r_start+2, column=c_start, value="TRANSFORM REQ:").font = Font(bold=True)
        ws.cell(row=r_start+2, column=c_start+1, value=instruction)
        
        # Varying table size for every PQ task
        rows = 5 + (tid % 8)
        cols = 3 + (tid % 3)
        ws.cell(row=r_start+4, column=c_start, value="DIRTY SOURCE").font = Font(bold=True)
        style_header(ws, r_start+5, c_start, cols, color=style['color'])
        for r in range(1, rows):
            for c in range(1, cols + 1):
                ws.cell(row=r_start+5+r, column=c_start+c-1, value=f"Data_{tid}_{r}_{c}")
        
        ws.cell(row=r_start+5, column=c_start+cols+1, value="CLEAN DESTINATION").font = Font(bold=True)
        for r in range(1, rows):
            ws.cell(row=r_start+5+r, column=c_start+cols+1).fill = PatternFill(start_color="FFFF00", fill_type="solid")
    return generate

def make_vba_gen(tid, instruction):
    def generate(ws):
        style = get_unique_style(tid)
        r_start = 2 + style['row_offset']
        c_start = 1 + style['col_offset']
        ws.title = f"VBA_{tid}"
        ws.cell(row=r_start, column=c_start, value=f"🏆 TASK {tid}: MACRO DEVELOPMENT").font = Font(size=14, bold=True, color=style['color'])
        ws.cell(row=r_start+2, column=c_start, value="SCRIPTING GOAL:").font = Font(bold=True)
        ws.cell(row=r_start+2, column=c_start+1, value=instruction)
        ws.cell(row=r_start+4, column=c_start, value="[ IDE EDITOR WINDOW ]").font = Font(bold=True)
        # VBA Box
        for r in range(r_start+5, r_start+18):
            for c in range(c_start, c_start+10):
                ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws.cell(row=r_start+5, column=c_start, value="' Sub Task_" + str(tid))
    return generate

# ============================================
# WORKBOOK ORCHESTRATION
# ============================================

def create_randomized_midterm(task_ids):
    from excel_assignment import create_excel_exercise_workbook
    wb = create_excel_exercise_workbook("Midterm_Final_Expert")
    
    if 'Instructions' not in wb.sheetnames:
        ws = wb.create_sheet("Instructions", 0)
    else:
        ws = wb['Instructions']
    
    # Absolute wipe of Instructions
    for r in range(1, 100):
        for c in range(1, 20):
            try:
                cell = ws.cell(row=r, column=c)
                if not isinstance(cell, openpyxl.cell.cell.MergedCell): cell.value = None
            except: pass

    # Absolute wipe of old sheets
    for sheetname in wb.sheetnames[:]:
        if sheetname != 'Instructions':
            try: wb.remove(wb[sheetname])
            except: pass
            
    bank = get_task_bank()
    for tid in task_ids:
        if tid in bank:
            # High unique sheet name
            # We must use a predictable pattern so the grader can find it
            new_ws = wb.create_sheet(title=f"Task_{tid}")
            bank[tid]['generate'](new_ws)
            
    ws['A1'] = "📊 RANDOMIZED MIDTERM EXAM - 100% UNIQUE & EXPERT"
    ws['A1'].font = Font(size=22, bold=True, color="C00000")
    ws['A3'] = "🚀 EXAMINATION INTEGRITY CHECK:"
    ws['A4'] = "• Your exam contains 10 unique sheets specifically chosen for you."
    ws['A5'] = "• Each sheet has a different layout, color, and technical challenge."
    ws['A7'] = f"Tasks Assigned: {', '.join(map(str, task_ids))}"
    
    ws.column_dimensions['A'].width = 80
    wb.active = ws
    return wb

def grade_randomized_midterm(file_path, task_ids):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except: return 0, "Error"
    bank = get_task_bank(); total_score = 0; details = []
    for tid in task_ids:
        if tid in bank:
            try:
                # Find the sheet for this task
                sheet_name = f"Task_{tid}"
                # Also check for old pattern if needed
                if sheet_name not in wb.sheetnames:
                    for s in wb.sheetnames:
                        if f"Exam_S{tid}_" in s:
                            sheet_name = s
                            break
                
                score = bank[tid]['grade'](wb, sheet_name)
                total_score += score
                details.append({'task': bank[tid]['title'], 'score': score})
            except Exception as e:
                print(f"Error grading Task {tid}: {e}")
                details.append({'task': bank[tid]['title'], 'score': 0})
    return total_score, details
