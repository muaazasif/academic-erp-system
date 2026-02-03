from app import app, db, Admin

print('Testing DB access...')
with app.app_context():
    try:
        admin = Admin.query.first()
        print('DB access successful:', 'Yes' if admin else 'No admin found')
        if admin:
            print(f'Found admin: {admin.username}')
    except Exception as e:
        print(f'DB access failed: {str(e)}')