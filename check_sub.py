import sqlite3

conn = sqlite3.connect('instance/bq-erp.db')
cursor = conn.cursor()
print(cursor.execute('SELECT id, student_id, assignment_id, score, allow_resubmit FROM excel_submission WHERE student_id = "3333"').fetchall())
conn.close()
