import psycopg2
from psycopg2 import sql

# Supabase Connection Config (Direct connection, port 5432)

DB_CONFIG = {
    "dbname": "postgres",
    "user": "postgres.ewuiehhszshxttnpmhjw",
    "password": "cANADA03248247884",
    "host": "aws-0-ap-southeast-1.pooler.supabase.com",
    "port": 5432
}

# Firebase Data
reviews_data = {
    "-OwMWS11RogQ35RFrs08": {
        "message": "Sir Muaaz Asif has been an excellent Data Analytics & Power BI instructor. He explains SQL and Excel concepts clearly, breaks down complex topics like joins, VLOOKUP, and dashboard building into manageable steps, and is patient with questions. His real-world experience in financial data and dashboards really comes through in how practically he teaches the material. I went from struggling with basics to feeling confident handling SQL queries and building Power BI dashboards. Highly recommend him anyone starting out in data analytics.",
        "name": "Muhammad Zaman ",
        "rating": "5",
        "school": "Escuela Schooling System",
        "timestamp": 1782806073715
    }
}

def import_reviews():
    try:
        print("Attempting direct connection to Supabase (port 5432)...")
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()
        print("✅ Connected to Supabase!")

        # Create table if not exists
        create_table_query = """
        CREATE TABLE IF NOT EXISTS reviews (
            firebase_id TEXT PRIMARY KEY,
            name TEXT,
            message TEXT,
            rating INTEGER,
            school TEXT,
            timestamp BIGINT
        );
        """
        cursor.execute(create_table_query)
        print("✅ Table 'reviews' ready.")

        # Insert data
        insert_query = """
        INSERT INTO reviews (firebase_id, name, message, rating, school, timestamp)
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (firebase_id) DO NOTHING;
        """
        
        for fb_id, data in reviews_data.items():
            cursor.execute(insert_query, (
                fb_id,
                data.get('name'),
                data.get('message'),
                int(data.get('rating', 0)),
                data.get('school'),
                data.get('timestamp')
            ))
            print(f"✅ Imported review: {fb_id}")

        conn.commit()
        cursor.close()
        conn.close()
        print("✅ Data import complete.")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    import_reviews()
