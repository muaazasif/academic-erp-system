import sqlite3

conn = sqlite3.connect('instance/bq-erp.db')
cursor = conn.cursor()

# Test global assignment allow_all_resubmits
cursor.execute('UPDATE excel_skills_assignment SET allow_all_resubmits = 1 WHERE id = 5')
conn.commit()
print("Assignment allow_all_resubmits:", cursor.execute('SELECT id, allow_all_resubmits FROM excel_skills_assignment WHERE id = 5').fetchall())

conn.close()
