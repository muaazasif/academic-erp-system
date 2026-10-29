from app import app, db
from sqlalchemy import text

with app.app_context():
    print("--- CHECKING ALL TABLES IN PUBLIC SCHEMA ---")
    res = db.session.execute(text("""
        SELECT table_name 
        FROM information_schema.tables 
        WHERE table_schema = 'public'
        ORDER BY table_name;
    """))
    tables = [row[0] for row in res]
    for t in tables:
        count_res = db.session.execute(text(f"SELECT count(*) FROM public.{t};"))
        print(f"Table: {t} -> Row count: {count_res.scalar()}")
