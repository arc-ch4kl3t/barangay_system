import mysql.connector
conn = mysql.connector.connect(host='localhost', user='root', password='', database='barangay_db')
cursor = conn.cursor(dictionary=True)

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
