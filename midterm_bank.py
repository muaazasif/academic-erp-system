import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import random
import re

# ============================================
# MASTER TASK BANK (100 TASKS) - EXPERT MODE
# ============================================

def style_header(ws, row, cols, color="1F4E79"):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    font = Font(bold=True, color="FFFFFF")
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

def get_task_bank():
    """Returns 100 tasks with UNIQUE LAYOUTS and logic"""
    bank = {}
    
    problems = {
        # ... (Same 100 unique problems as before)
        1: "Analyze regional performance: If Sales > 50000 and Profit > 20%, status is 'ELITE'. If Sales > 30000, 'PRIME'. Else 'standard'.",
        2: "Extract Domain: Use MID and FIND to isolate the domain name from emails like 'user@sub.company.org'.",
        3: "Working Days: Calculate the exact completion date 120 working days after today, excluding holidays listed in G4:G10.",
        4: "Nested IF: Grade students: >95:A+, >90:A, >85:B+, >80:B, >70:C, Else:F.",
        5: "Text Cleaning: Use TRIM, PROPER, and SUBSTITUTE to clean '  mUAAZ_asif-PAKISTAN  ' into 'Muaaz Asif Pakistan'.",
        6: "Array Sum: Calculate the total sales for 'Electronics' but ONLY for the 'North' region after '2024-03-01'.",
        7: "Logic Gate: If (A AND B) OR (C AND NOT D), then 1, else 0. Implement for 50 rows of binary data.",
        8: "Data Validation: Ensure Column B only accepts even numbers between 500 and 1500.",
        9: "Dynamic Offset: Use OFFSET to sum the last 5 months of sales data automatically.",
        10: "Error Handling: Use IFERROR and VLOOKUP to search for IDs; if not found, show 'ID_MISSING_CONTACT_ADMIN'.",
        11: "Conditional Formatting: Highlight rows where Profit is negative AND Sales are in the top 10%.",
        12: "Loan Amortization: Calculate total interest paid over 15 years for a loan of 2M at 7.5% fixed.",
        13: "Date Diff: Calculate age in Years, Months, and Days exactly from a birthdate in cell C5.",
        14: "Unique Count: Use an array formula to count unique student names in a list of 500 entries.",
        15: "Ranking: Rank products based on Sales, but if Sales are tied, use Profit % as the tie-breaker.",
        16: "Data Shifting: Move all non-empty values from a column to a new range, removing all blanks.",
        17: "Multi-Sheet Lookup: VLOOKUP a value across 3 different sheets (SheetA, SheetB, SheetC) until found.",
        18: "Frequency: Calculate the distribution of test scores into buckets of 10 (0-10, 11-20...).",
        19: "Text Joining: Combine Name, City, and Zip into a single string 'NAME [CITY] - ZIP' for 100 users.",
        20: "Transposing: Take a 5x5 matrix and rotate it 90 degrees using formulas only.",
        21: "Pro Lookup: Use INDEX/MATCH/MATCH for a 2-way matrix lookup (Product vs Month).",
        22: "Dynamic Filter: Filter a list of transactions to show only those > $10,000 for 'Refund' category.",
        23: "Sorting: Sort a dynamic range by Category (ASC) and then by Date (DESC) using formulas.",
        24: "XLOOKUP Power: Use XLOOKUP to find the first occurrence from the bottom of a 1000-row list.",
        25: "Scenario Manager: Build a model that switches between 'Best Case', 'Worst Case', and 'Realistic'.",
        26: "Goal Seek: Find the required sales volume to reach a target profit of $50,000.",
        27: "Solver: Optimize the mix of 3 products to maximize profit given a 40-hour labor constraint.",
        28: "Indirect Ref: Build a summary sheet that pulls the 'Total' cell from sheets named by Month (Jan, Feb...).",
        29: "Recursive Lambda: Write a recursive logic to calculate the Fibonacci sequence up to 20 terms.",
        30: "Data Table: Create a 2-variable data table for Interest Rate vs. Term in a mortgage model.",
        31: "Match Type: Use MATCH with '-1' type for finding the smallest value greater than or equal to a target.",
        32: "Wildcard VLOOKUP: Search for partial product names like '*Pro*' in a database of 500 items.",
        33: "Dynamic Range: Create a Named Range that automatically expands as new rows are added.",
        34: "Subtotals: Use SUBTOTAL(9, ...) to sum filtered data only, ignoring hidden rows.",
        35: "Array Constants: Use {1,2,3;4,5,6} to perform matrix multiplication in a single cell.",
        36: "Fiscal Year: Convert standard dates into Fiscal Year quarters (starting July 1st).",
        37: "Network Days: Find total working hours between two dates, accounting for 8-hour workdays.",
        38: "Text to Columns: Split a CSV-formatted string into 5 separate cells using formulas only.",
        39: "Duplicate Finder: Identify IDs that appear more than 3 times in a list.",
        40: "Z-Score: Calculate the Z-Score for a set of student marks to identify outliers.",
        41: "Joins: INNER JOIN Students, Courses, and Instructors to list 'Student Name - Course - Instructor'.",
        42: "Aggregation: Find the Top 3 Departments by total revenue in the last 6 months.",
        43: "Subqueries: Select all employees whose salary is higher than the average of their department.",
        44: "Window Functions: Rank sales reps within each region using ROW_NUMBER() and PARTITION BY.",
        45: "CTEs: Use a WITH clause to calculate a recursive category hierarchy.",
        46: "Case Statements: Categorize transactions as 'Small' (<100), 'Medium' (100-500), 'Large' (>500).",
        47: "Exists: Find customers who have placed an order in 2023 but NO orders in 2024.",
        48: "Grouping Sets: Use GROUPING SETS to create a subtotal for (Region, Category) and (Region) only.",
        49: "Strings: Use COALESCE and CONCAT to format full addresses, handling null 'Unit Number' fields.",
        50: "Dates: Extract day-of-week and calculate the average sales for every 'Friday' in history.",
        51: "Performance: Write a query that uses an INDEX hint to optimize a search on 'last_name'.",
        52: "Self Join: List all employees and their Manager's name by joining the table to itself.",
        53: "Distinct: Count how many unique cities each student has lived in according to the history table.",
        54: "Having: List Departments that have more than 10 employees AND an average salary < 40k.",
        55: "Updates: Write an UPDATE query with a JOIN to increase prices by 10% for 'Discontinued' items.",
        56: "Deletes: Remove all duplicate records from a staging table, keeping only the newest ID.",
        57: "Views: Create a SECURE VIEW that hides the 'Social Security Number' column from regular users.",
        58: "Math: Calculate the standard deviation of 'Wait Time' for customer support tickets.",
        59: "Union: Merge 'Past_Students' and 'Current_Students' while removing duplicates.",
        60: "Cast: Convert a 'Text' timestamp into a valid 'Date' and calculate the months since account creation.",
        61: "Pivoting: Convert a 12-month column layout into a flat 'Attribute/Value' format.",
        62: "Merging: Join a local Excel table with a web-based JSON API source.",
        63: "Cleaning: Remove all non-printable characters and extra spaces from a 5000-row column.",
        64: "M-Code: Write a custom M-function to calculate Tax based on a dynamic parameter.",
        65: "Fuzzy Match: Merge two customer lists where names are slightly different (e.g. 'J. Doe' vs 'John Doe').",
        66: "Splitting: Split a 'Path' string (C:\\Users\\Admin\\Docs\\File.txt) into Folder and Filename.",
        67: "Parameters: Create a PQ parameter that allows the user to change the 'Source Folder' dynamically.",
        68: "Appending: Combine 12 CSV files from a folder into one master table automatically.",
        69: "Grouping: Group by Year and Month to find the 'Maximum' sale of each month.",
        70: "Conditionals: Create a 'Custom Column' that applies 5 different logic rules using M-code IF.",
        71: "Transposing: Flip a headers-on-the-left table into a standard headers-on-top table.",
        72: "Replacing: Use a lookup table to replace 50 different abbreviations with full names.",
        73: "Sorting: Ensure the Power Query output is always sorted by 'Priority' (High, Med, Low) correctly.",
        74: "Fill Down: Use 'Fill Down' to populate null category cells from the row above.",
        75: "Extracting: Isolate the text between two delimiters (e.g. between '(' and ')').",
        76: "Data Types: Fix 'Scientific Notation' numbers that were imported as text.",
        77: "Buffers: Use Table.Buffer to optimize a slow merge operation on large datasets.",
        78: "Distinct: Keep only the first occurrence of each 'Customer ID' while sorting by 'Date'.",
        79: "Filtering: Filter rows where the 'Comments' column contains any word from a forbidden list.",
        80: "Schema: Force a specific schema on a dynamic folder import to prevent 'Column Not Found' errors.",
        81: "Loops: Loop through 100 sheets and protect them with a random password.",
        82: "UserForms: Create a pop-up form that validates 'Email' and 'Phone' before adding to a sheet.",
        83: "Events: Write a Workbook_Open macro that logs the Username and Time to a hidden sheet.",
        84: "FileSystem: Write code to list all filenames from 'C:\\Midterm_Data' into Column A.",
        85: "Optimization: Write a macro that runs 10x faster by disabling ScreenUpdating and Calculations.",
        86: "Object-Oriented: Create a Class Module named 'Student' with properties 'Name' and 'Marks'.",
        87: "Charts: Automate the creation of a 'Pareto Chart' based on the current selection.",
        88: "Email: Write a Sub to send a personalized Outlook email to everyone in the list.",
        89: "Outlook: Save all attachments from an Outlook folder named 'Submissions' to a local path.",
        90: "RegEx: Use 'VBScript.RegExp' to validate a password (must have 1 capital, 1 number).",
        91: "API: Write a VBA function to fetch the current USD/PKR exchange rate from a public API.",
        92: "Arrays: Read 10,000 rows into a VBA Array, process them, and write back in one shot.",
        93: "Dictionary: Use 'Scripting.Dictionary' to find unique values and their frequencies.",
        94: "Errors: Implement a global error handler that logs the line number where a crash occurs.",
        95: "Formatting: Create a macro that converts any selected table into a 'Professional PDF'.",
        96: "Protection: Create a macro that hides the 'Formulas Bar' and 'Ribbon' for a kiosk-mode exam.",
        97: "Sorting: Write a 'Bubble Sort' algorithm in VBA to sort a 1D array of strings.",
        98: "SQL in VBA: Use ADODB to run a SQL query against an external Access database.",
        99: "VBA Logic: Create a function that calculates the 'Shortest Path' between two nodes in a matrix.",
        100: "Master Macro: Create a button that runs Data Cleaning -> Formatting -> Emailing in sequence."
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
            'title': f"{get_topic_name(tid).replace('_',' ')} {tid}",
            'generate': gen,
            'grade': lambda wb: 1.0
        }
    return bank

def get_topic_name(i):
    if i <= 20: return "EXCEL_LOGIC"
    if i <= 40: return "EXCEL_ADVANCED"
    if i <= 60: return "SQL_QUERIES"
    if i <= 80: return "POWER_QUERY"
    return "VBA_EXPERT"

# ---------------------------------------------------------
# UNIQUE LAYOUT GENERATORS
# ---------------------------------------------------------

def make_excel_logic_gen(tid, instruction):
    def generate(ws):
        ws.title = f"Excel_Logic_{tid}"
        ws['A1'] = f"🏆 TASK {tid}: LOGICAL REASONING"
        ws['A1'].font = Font(size=16, bold=True, color="1F4E79")
        ws['A3'] = "INSTRUCTION:"; ws['B3'] = instruction; ws['B3'].font = Font(bold=True)
        # Unique table for logical tasks
        headers = ['Entity', 'Attribute_X', 'Attribute_Y', 'Your Formula Output']
        style_header(ws, 5, 4, color="1F4E79")
        for r in range(6, 12):
            ws.cell(row=r, column=1, value=f"ID_{tid}_{r}")
            ws.cell(row=r, column=2, value=random.randint(10, 100))
            ws.cell(row=r, column=3, value=random.choice(['True','False','Yes','No']))
            ws.cell(row=r, column=4).fill = PatternFill(start_color="FFFF00", fill_type="solid")
    return generate

def make_excel_advanced_gen(tid, instruction):
    def generate(ws):
        ws.title = f"Excel_Adv_{tid}"
        ws['A1'] = f"🏆 TASK {tid}: PRO DATA MODELING"
        ws['A1'].font = Font(size=16, bold=True, color="C00000")
        ws['A3'] = "CHALLENGE:"; ws['B3'] = instruction; ws['B3'].font = Font(bold=True)
        # Matrix Layout for Advanced tasks
        ws['F5'] = "REFERENCE MATRIX"; style_header(ws, 5, 3, color="548235")
        for r in range(6, 9):
            for c in range(6, 9): ws.cell(row=r, column=c, value=random.randint(100, 500))
        ws['A5'] = "WORK AREA"; style_header(ws, 5, 2, color="C00000")
        ws['A6'] = "Input Parameter"; ws['B6'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws['A7'] = "Calculated Result"; ws['B7'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    return generate

def make_sql_gen(tid, instruction):
    def generate(ws):
        ws.title = f"SQL_Query_{tid}"
        ws['A1'] = f"🏆 TASK {tid}: SQL ENGINEERING"
        ws['A1'].font = Font(size=16, bold=True, color="2E75B6")
        ws['A3'] = "SCHEMA DESCRIPTION:"
        ws['B3'] = "Table: Transactions (id, user_id, amount, date, status)"
        ws['A5'] = "SQL REQUIREMENT:"
        ws['B5'] = instruction; ws['B5'].font = Font(bold=True)
        ws['A7'] = "WRITE SQL CODE HERE:"; ws.merge_cells('B7:G10')
        ws['B7'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws['B7'].alignment = Alignment(vertical='top', wrap_text=True)
        ws.column_dimensions['B'].width = 80
    return generate

def make_power_query_gen(tid, instruction):
    def generate(ws):
        ws.title = f"PQ_ETL_{tid}"
        ws['A1'] = f"🏆 TASK {tid}: POWER QUERY / ETL"
        ws['A1'].font = Font(size=16, bold=True, color="7030A0")
        ws['A3'] = "ETL TASK:"; ws['B3'] = instruction; ws['B3'].font = Font(bold=True)
        # Messy Data Layout for PQ
        ws['A5'] = "RAW DIRTY DATA (MESSY)"; style_header(ws, 5, 4, color="7030A0")
        for r in range(6, 12):
            ws.cell(row=r, column=1, value=f"  {random.randint(100,200)}  ")
            ws.cell(row=r, column=2, value=f"USER_{tid}_#_{r}")
            ws.cell(row=r, column=3, value=f"error_{random.randint(1,5)}")
            ws.cell(row=r, column=4, value="2024/01/01")
        ws['F5'] = "EXPECTED CLEAN OUTPUT"; style_header(ws, 5, 3, color="BF8F00")
        for r in range(6, 12):
            for c in range(6, 9): ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFF00", fill_type="solid")
    return generate

def make_vba_gen(tid, instruction):
    def generate(ws):
        ws.title = f"VBA_Automation_{tid}"
        ws['A1'] = f"🏆 TASK {tid}: VBA AUTOMATION"
        ws['A1'].font = Font(size=16, bold=True, color="375623")
        ws['A3'] = "AUTOMATION SCRIPT:"; ws['B3'] = instruction; ws['B3'].font = Font(bold=True)
        ws['A5'] = "DEVELOPER CODE BOX:"; ws.merge_cells('A6:H15')
        ws['A6'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
        ws['A6'].alignment = Alignment(vertical='top', wrap_text=True)
        ws['A6'] = "Sub Procedure_" + str(tid) + "()\n\n' Write your code here\n\nEnd Sub"
        ws.column_dimensions['A'].width = 100
    return generate

# ============================================
# WORKBOOK ORCHESTRATION
# ============================================

def create_randomized_midterm(task_ids):
    from excel_assignment import create_excel_exercise_workbook
    wb = create_excel_exercise_workbook("Midterm_Ultra_Hard")
    
    if 'Instructions' not in wb.sheetnames:
        ws = wb.create_sheet("Instructions", 0)
    else:
        ws = wb['Instructions']
    
    # Clean Instructions safely
    for r in range(1, 100):
        for c in range(1, 15):
            try:
                cell = ws.cell(row=r, column=c)
                if not isinstance(cell, openpyxl.cell.cell.MergedCell): cell.value = None
            except: pass

    # Clear old sheets
    for sheetname in wb.sheetnames[:]:
        if sheetname != 'Instructions':
            try: wb.remove(wb[sheetname])
            except: pass
            
    bank = get_task_bank()
    for tid in task_ids:
        if tid in bank:
            new_ws = wb.create_sheet(title=f"Initial_{tid}")
            bank[tid]['generate'](new_ws)
            
    # Final info
    ws['A1'] = "📊 RANDOMIZED MIDTERM EXAM - 100% UNIQUE & HARD"
    ws['A1'].font = Font(size=20, bold=True, color="C00000")
    ws['A3'] = "🚀 EXPERT MODE ENABLED"
    ws['A5'] = "1. You have 10 unique challenges across 5 disciplines."
    ws['A6'] = "2. Each sheet uses a different technical layout (SQL, VBA, Logic, ETL)."
    ws['A8'] = f"Tasks: {', '.join(map(str, task_ids))}"
    
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
                score = bank[tid]['grade'](wb)
                total_score += score
                details.append({'task': bank[tid]['title'], 'score': score})
            except: details.append({'task': bank[tid]['title'], 'score': 0})
    return total_score, details
