#!/usr/bin/env python
"""Fix user roles - set all to admin"""

import mysql.connector

conn = mysql.connector.connect(
    host='localhost',
    user='root',
    password='',
    database='barangay_db'
)
cursor = conn.cursor(dictionary=True)

# Update all users to admin role
print("Updating user roles to 'admin'...")
cursor.execute("UPDATE users SET role = 'admin' WHERE role != 'admin'")
conn.commit()

# Verify
cursor.execute('SELECT username, role FROM users;')
users = cursor.fetchall()

print('\n✓ Users updated:')
for user in users:
    print(f"  {user['username']}: role={user['role']}")

conn.close()
print('\n✅ All users now have admin role!')
