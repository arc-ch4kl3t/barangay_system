#!/usr/bin/env python
"""Fix user roles - set all to admin"""

import os
import psycopg2
from psycopg2.extras import RealDictCursor

database_url = os.environ.get("DATABASE_URL")
if not database_url:
    print("❌ DATABASE_URL environment variable not set!")
    exit(1)

conn = psycopg2.connect(database_url, sslmode="require")
cursor = conn.cursor(cursor_factory=RealDictCursor)

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
