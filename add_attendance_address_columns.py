from app import app, db
with app.app_context():
    with db.engine.connect() as conn:
        try:
            conn.execute(db.text("ALTER TABLE attendance ADD COLUMN IF NOT EXISTS check_in_address TEXT;"))
            conn.execute(db.text("ALTER TABLE attendance ADD COLUMN IF NOT EXISTS check_out_address TEXT;"))
            conn.commit()
            print("✅ Added check_in_address and check_out_address columns to attendance table successfully!")
        except Exception as e:
            print(f"Note/Error: {e}")
