import re

with open('replace_excel_submissions_383.py', 'r', encoding='utf-8') as f:
    c1 = f.read()
with open('restore_excel_submissions.py', 'r', encoding='utf-8') as f:
    c2 = f.read()

m1 = re.search(r'raw_data\s*=\s*"""(.*?)"""', c1, re.DOTALL)
m2 = re.search(r'raw_data\s*=\s*"""(.*?)"""', c2, re.DOTALL)

data1 = m1.group(1) if m1 else ''
data2 = m2.group(1) if m2 else ''

def parse_data(raw):
    lines = raw.strip().split('\n')
    records = []
    current_record = None
    for line in lines:
        if not line.strip():
            continue
        parts = line.split('\t')
        if len(parts) >= 6 and parts[0].strip().isdigit() and len(parts[0].strip()) < 10:
            if current_record:
                records.append(current_record)
            current_record = parts
        else:
            if current_record:
                current_record[-1] += "\n" + line
    if current_record:
        records.append(current_record)
    return records

rec1 = parse_data(data1)
rec2 = parse_data(data2)
print("Records in file 1:", len(rec1))
print("Records in file 2:", len(rec2))
