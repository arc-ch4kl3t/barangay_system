import mysql.connector

# Read the database.sql file
with open('database.sql', 'r') as f:
    sql_content = f.read()

# Connect to MySQL
conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password=""
)

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
