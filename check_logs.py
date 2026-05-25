import os
import psycopg2
from psycopg2.extras import RealDictCursor

database_url = os.environ.get("DATABASE_URL")
if not database_url:
    print("❌ DATABASE_URL environment variable not set!")
    exit(1)

conn = psycopg2.connect(database_url, sslmode="require")
cursor = conn.cursor(cursor_factory=RealDictCursor)

# Check audit_logs table data
print('=== AUDIT LOGS DATA ===')
cursor.execute('SELECT * FROM audit_logs ORDER BY created_at DESC')
logs = cursor.fetchall()
print(f'Total logs: {len(logs)}')
print()
for log in logs:
    print(f'ID: {log.get("id")}')
    print(f'Username: {log.get("username")}')
    print(f'Action Type: {log.get("action_type")}')
    print(f'Details: {log.get("details")}')
    print(f'Household Context: {log.get("household_context")}')
    print(f'Created At: {log.get("created_at")}')
    print('---')

conn.close()
