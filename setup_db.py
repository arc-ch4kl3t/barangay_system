import os
import psycopg2

# Read the database.sql file
with open('database.sql', 'r') as f:
    sql_content = f.read()

# Connect to PostgreSQL
database_url = os.environ.get("DATABASE_URL")
if not database_url:
    print("❌ DATABASE_URL environment variable not set!")
    exit(1)

conn = psycopg2.connect(database_url, sslmode="require")
cursor = conn.cursor()

# Execute the SQL statements
for statement in sql_content.split(';'):
    statement = statement.strip()
    if statement:
        try:
            cursor.execute(statement)
            print(f"✓ Executed: {statement[:60]}...")
        except Exception as e:
            print(f"✗ Error executing statement: {e}")

conn.commit()
cursor.close()
conn.close()
print("\nDatabase setup complete!")
