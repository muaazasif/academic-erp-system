"""
Migrate from SQLite to PostgreSQL (Supabase/Neon/Render free tier)
This ensures your data is permanent and won't be deleted
"""
import os
import sqlite3
import psycopg2
from psycopg2.extensions import ISOLATION_LEVEL_AUTOCOMMIT

# SQLite source
SQLITE_DB = 'instance/erp_system.db'

# PostgreSQL destination - UPDATE THESE WITH YOUR FREE PROVIDER DETAILS
# Free options:
# 1. Supabase (supabase.com) - 500MB free
# 2. Neon (neon.tech) - 500MB free
# 3. Render (render.com) - 90 days free
# 4. CockroachDB (cockroachlabs.cloud) - 5GB free forever
PG_CONFIG = {
    'dbname': 'your_database',
    'user': 'your_username',
    'password': 'your_password',
    'host': 'your_host.database.example.com',
    'port': '5432'
}

def get_sqlite_connection():
    """Connect to SQLite database"""
    if not os.path.exists(SQLITE_DB):
        raise FileNotFoundError(f"SQLite database not found: {SQLITE_DB}")
    return sqlite3.connect(SQLITE_DB)

def get_postgres_connection(dbname='postgres'):
    """Connect to PostgreSQL server"""
    config = PG_CONFIG.copy()
    if dbname != 'postgres':
        config['dbname'] = dbname
    return psycopg2.connect(**config)

def create_target_database():
    """Create the target database if it doesn't exist"""
    try:
        # Connect to default postgres database
        conn = get_postgres_connection('postgres')
        conn.set_isolation_level(ISOLATION_LEVEL_AUTOCOMMIT)
        cursor = conn.cursor()
        
        # Check if database exists
        cursor.execute("SELECT 1 FROM pg_database WHERE datname = %s", (PG_CONFIG['dbname'],))
        if not cursor.fetchone():
            cursor.execute(f"CREATE DATABASE {PG_CONFIG['dbname']}")
            print(f"✅ Created database: {PG_CONFIG['dbname']}")
        else:
            print(f"ℹ️ Database already exists: {PG_CONFIG['dbname']}")
        
        cursor.close()
        conn.close()
        return True
    except Exception as e:
        print(f"❌ Error creating database: {e}")
        return False

def get_sqlite_tables(sqlite_conn):
    """Get all tables from SQLite"""
    cursor = sqlite_conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    return [row[0] for row in cursor.fetchall()]

def get_table_schema(sqlite_conn, table_name):
    """Get table schema from SQLite"""
    cursor = sqlite_conn.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    return cursor.fetchall()

def sqlite_to_postgres_type(sqlite_type):
    """Convert SQLite types to PostgreSQL types"""
    type_mapping = {
        'INTEGER': 'INTEGER',
        'TEXT': 'TEXT',
        'REAL': 'REAL',
        'BLOB': 'BYTEA',
        'NUMERIC': 'NUMERIC',
        'BOOLEAN': 'BOOLEAN',
        'DATETIME': 'TIMESTAMP',
        'VARCHAR': 'VARCHAR'
    }
    return type_mapping.get(sqlite_type.upper(), 'TEXT')

def create_table_in_postgres(pg_conn, table_name, schema):
    """Create table in PostgreSQL"""
    cursor = pg_conn.cursor()
    
    columns = []
    for col in schema:
        col_id, col_name, col_type, notnull, default_value, is_pk = col
        
        # Convert type
        pg_type = sqlite_to_postgres_type(col_type)
        
        # Build column definition
        col_def = f"{col_name} {pg_type}"
        
        if is_pk:
            col_def += " PRIMARY KEY"
        
        if notnull and not is_pk:
            col_def += " NOT NULL"
        
        if default_value is not None:
            col_def += f" DEFAULT {default_value}"
        
        columns.append(col_def)
    
    # Handle AUTOINCREMENT for INTEGER PRIMARY KEY
    has_autoincrement = any(col[5] and col[2].upper() == 'INTEGER' for col in schema)
    
    columns_str = ', '.join(columns)
    
    # Drop table if exists (for migration)
    cursor.execute(f"DROP TABLE IF EXISTS {table_name}")
    
    create_sql = f"CREATE TABLE {table_name} ({columns_str})"
    cursor.execute(create_sql)
    
    pg_conn.commit()
    print(f"✅ Created table: {table_name}")

def migrate_table_data(sqlite_conn, pg_conn, table_name):
    """Migrate data from SQLite to PostgreSQL"""
    sqlite_cursor = sqlite_conn.cursor()
    pg_cursor = pg_conn.cursor()
    
    # Get all data from SQLite
    sqlite_cursor.execute(f"SELECT * FROM {table_name}")
    rows = sqlite_cursor.fetchall()
    
    if not rows:
        print(f"ℹ️ Table {table_name} is empty, skipping data migration")
        return 0
    
    # Get column names
    sqlite_cursor.execute(f"PRAGMA table_info({table_name})")
    columns = [col[1] for col in sqlite_cursor.fetchall()]
    
    # Prepare INSERT statement
    placeholders = ', '.join(['%s'] * len(columns))
    columns_str = ', '.join(columns)
    insert_sql = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
    
    # Insert data
    inserted = 0
    for row in rows:
        try:
            # Convert None to proper NULL
            clean_row = tuple(None if val == 'None' else val for val in row)
            pg_cursor.execute(insert_sql, clean_row)
            inserted += 1
        except Exception as e:
            print(f"  ⚠️ Error inserting row: {e}")
    
    pg_conn.commit()
    print(f"✅ Migrated {inserted}/{len(rows)} rows from {table_name}")
    return inserted

def verify_migration(sqlite_conn, pg_conn, table_name):
    """Verify data was migrated correctly"""
    sqlite_cursor = sqlite_conn.cursor()
    pg_cursor = pg_conn.cursor()
    
    sqlite_cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
    sqlite_count = sqlite_cursor.fetchone()[0]
    
    pg_cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
    pg_count = pg_cursor.fetchone()[0]
    
    if sqlite_count == pg_count:
        print(f"✅ Verified: {table_name} has {pg_count} rows in both databases")
        return True
    else:
        print(f"❌ MISMATCH: {table_name} - SQLite: {sqlite_count}, PostgreSQL: {pg_count}")
        return False

def migrate_all():
    """Main migration function"""
    print("=" * 70)
    print("🔄 SQLITE TO POSTGRESQL MIGRATION TOOL")
    print("=" * 70)
    print(f"\nSource: {SQLITE_DB}")
    print(f"Target: {PG_CONFIG['host']}/{PG_CONFIG['dbname']}")
    print()
    
    # Test connections
    try:
        sqlite_conn = get_sqlite_connection()
        print("✅ Connected to SQLite")
    except Exception as e:
        print(f"❌ Failed to connect to SQLite: {e}")
        return False
    
    try:
        pg_conn = get_postgres_connection(PG_CONFIG['dbname'])
        print("✅ Connected to PostgreSQL")
    except Exception as e:
        print(f"❌ Failed to connect to PostgreSQL: {e}")
        print("\n💡 UPDATE the PG_CONFIG in this file with your database credentials!")
        print("\nFree PostgreSQL options:")
        print("  1. Supabase: https://supabase.com (500MB free)")
        print("  2. Neon: https://neon.tech (500MB free)")
        print("  3. CockroachDB: https://cockroachlabs.cloud (5GB free)")
        return False
    
    # Create database if needed
    create_target_database()
    
    # Get tables
    tables = get_sqlite_tables(sqlite_conn)
    print(f"\n📋 Found {len(tables)} tables to migrate: {', '.join(tables)}")
    
    # Migrate each table
    total_rows = 0
    for table_name in tables:
        schema = get_table_schema(sqlite_conn, table_name)
        create_table_in_postgres(pg_conn, table_name, schema)
        rows = migrate_table_data(sqlite_conn, pg_conn, table_name)
        total_rows += rows
    
    # Verify
    print("\n" + "=" * 70)
    print("🔍 Verifying migration...")
    print("=" * 70)
    all_verified = True
    for table_name in tables:
        if not verify_migration(sqlite_conn, pg_conn, table_name):
            all_verified = False
    
    # Summary
    print("\n" + "=" * 70)
    if all_verified:
        print("✅ MIGRATION COMPLETE!")
        print(f"📊 Total rows migrated: {total_rows}")
        print(f"💾 Your data is now safely stored in PostgreSQL")
        print(f"🎯 This data will NOT be deleted when Railway expires!")
    else:
        print("⚠️ Migration completed with some verification issues")
        print("💡 Please review the errors above")
    print("=" * 70)
    
    sqlite_conn.close()
    pg_conn.close()
    
    return all_verified

if __name__ == '__main__':
    migrate_all()
