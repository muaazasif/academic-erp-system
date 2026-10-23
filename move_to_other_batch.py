import os
import sys
from app import app, db, Student, Batch

def move_dummies():
    with app.app_context():
        # 1. Ensure 'Other' batch exists
        other_batch = Batch.query.filter_by(name='Other').first()
        if not other_batch:
            other_batch = Batch(name='Other', is_active=False)
            db.session.add(other_batch)
            db.session.commit()
            print("✅ Created 'Other' batch.")
        else:
            print("✅ 'Other' batch already exists.")

        # Dummy IDs and Names list provided by user
        dummy_identifiers = [
            '8878', '123', '536400', '4646', '1234', '5035755', '493939', '4919811', 
            '50416888', '1999', '7777', '101', 'Mahek Naseem', '50543', '503818', 
            '3333', '98989', '7888', '4747', 'TEST', 'TEST_CLASS', 'CS101', 'MATH201', 
            'MATH101', 'TEST001', 'TEST_FAILED_SYNC', 'TEST123', 'ayman', 'muaaz', 'muz', 
            'humnah', 'ghazi', 'sidra', '499448', '4754111', '12345', '475411111', 
            'admin', '5041688', '4969522', '122', '5046222', '50462222', '2000', 
            '2001', '2003', '2010', '1444', '1555', '1222', '1111', '6666', '3874522', 
            '5073122', '9991', '9992', '9993', '9994', '9995', '9999', '99900', 
            '9900', '9000', '9898', '7878', '1968', '9891', '5656', '5555', '8080', 
            '8888', '47541111', '50357555', '3431844', '4932400', '2007', '444', 
            '4444', '5054933'
        ]

        dummy_set_lower = {str(x).strip().lower() for x in dummy_identifiers}

        updated_count = 0
        for s in Student.query.all():
            s_id = str(s.student_id).strip().lower()
            s_name = str(s.name).strip().lower()
            if s_id in dummy_set_lower or s_name in dummy_set_lower:
                s.batch = 'Other'
                updated_count += 1

        db.session.commit()
        print(f"✅ Successfully moved {updated_count} dummy students to 'Other' batch!")

if __name__ == "__main__":
    move_dummies()
