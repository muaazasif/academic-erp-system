import unittest
from app import app, db, Admin, Student, ExcelSkillsAssignment, ExcelAssignment

class ExcelAccessTestCase(unittest.TestCase):
    def setUp(self):
        app.config['TESTING'] = True
        app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
        self.client = app.test_client()
        with app.app_context():
            db.create_all()
            # Create test admin
            admin = Admin(username='testadmin')
            admin.set_password('admin123')
            db.session.add(admin)
            
            # Create test student
            student = Student(student_id='ST101', name='Test Student')
            student.set_password('pass123')
            db.session.add(student)
            
            # Create test excel assignment
            excel_assign = ExcelSkillsAssignment(title='Excel Skill 1 VLOOKUP', description='Test VLOOKUP')
            db.session.add(excel_assign)
            db.session.commit()
            self.excel_assign_id = excel_assign.id

    def tearDown(self):
        with app.app_context():
            db.drop_all()

    def test_admin_assign_route(self):
        with self.client as c:
            # Login as admin
            with c.session_transaction() as sess:
                sess['admin_id'] = 1
                sess['admin_username'] = 'testadmin'
            
            # Test GET assign page
            response = c.get(f'/admin/excel-assignments/{self.excel_assign_id}/assign')
            self.assertEqual(response.status_code, 200)
            self.assertIn(b'Excel Skill 1 VLOOKUP', response.data)
            self.assertIn(b'Yes', response.data)
            self.assertIn(b'No', response.data)
            
            # Test POST assign page (assigning ST101 with Yes)
            response = c.post(f'/admin/excel-assignments/{self.excel_assign_id}/assign', data={
                'access_ST101': 'yes'
            }, follow_redirects=True)
            self.assertEqual(response.status_code, 200)
            
            # Verify in DB that ExcelAssignment was created for ST101
            with app.app_context():
                record = ExcelAssignment.query.filter_by(assignment_id=self.excel_assign_id, student_id='ST101').first()
                self.assertIsNotNone(record)

    def test_student_access_control(self):
        with self.client as c:
            # Login as student without assignment access
            with c.session_transaction() as sess:
                sess['student_id'] = 'ST101'
            
            response = c.get('/student/excel-assignments')
            self.assertEqual(response.status_code, 200)
            # Should not see the assignment since not assigned yet
            self.assertNotIn(b'Excel Skill 1 VLOOKUP', response.data)
            
            # Now assign via admin
            with app.app_context():
                ea = ExcelAssignment(assignment_id=self.excel_assign_id, student_id='ST101')
                db.session.add(ea)
                db.session.commit()
            
            # Check student view again
            response = c.get('/student/excel-assignments')
            self.assertEqual(response.status_code, 200)
            self.assertIn(b'Excel Skill 1 VLOOKUP', response.data)

if __name__ == '__main__':
    unittest.main()
