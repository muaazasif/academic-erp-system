from app import app, db
with app.app_context():
    with db.engine.connect() as conn:
        try:
            conn.execute(db.text("ALTER TABLE student ADD COLUMN IF NOT EXISTS gender VARCHAR(50) DEFAULT 'Female';"))
            conn.commit()
            print("✅ Added gender column to student table successfully!")
        except Exception as e:
            print(f"Note/Error: {e}")
