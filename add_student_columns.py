from app import app, db
with app.app_context():
    with db.engine.connect() as conn:
        try:
            conn.execute(db.text("ALTER TABLE student ADD COLUMN IF NOT EXISTS batch VARCHAR(150) DEFAULT 'Bano Qabil 5.0 Fall';"))
            conn.execute(db.text("ALTER TABLE student ADD COLUMN IF NOT EXISTS campus VARCHAR(250);"))
            conn.execute(db.text("ALTER TABLE student ADD COLUMN IF NOT EXISTS timings VARCHAR(100);"))
            conn.commit()
            print("✅ Added batch, campus, timings columns to student table successfully!")
        except Exception as e:
            print(f"Note/Error: {e}")
