#!/usr/bin/env python
"""Verify Step 1: Database Migration"""

import mysql.connector
import sys

try:
    conn = mysql.connector.connect(
        host='localhost',
        user='root',
        password='',
        database='barangay_db'
    )
    cursor = conn.cursor(dictionary=True)

    print("\n" + "="*50)
    print("STEP 1 VERIFICATION - Database Migration")
    print("="*50)

    # Check 1: Users table has role column
    print("\n✓ CHECKING USERS TABLE...")
    cursor.execute('DESCRIBE users;')
    columns = cursor.fetchall()
    has_role = False
    for col in columns:
        if col['Field'] == 'role':
            has_role = True
            print(f"  ✓ role column found: {col['Type']}")
        if col['Field'] == 'username':
            print(f"  ✓ username column: {col['Type']}")
        if col['Field'] == 'password':
            print(f"  ✓ password column: {col['Type']}")

    if not has_role:
        print("  ✗ ERROR: role column NOT found!")
        sys.exit(1)

    # Check 2: Password resets table exists
    print("\n✓ CHECKING PASSWORD_RESETS TABLE...")
    try:
        cursor.execute('DESCRIBE password_resets;')
        columns = cursor.fetchall()
        print("  ✓ Table exists with columns:")
        for col in columns:
            print(f"    - {col['Field']}: {col['Type']}")
    except Exception as e:
        print(f"  ✗ ERROR: Table does not exist: {e}")
        sys.exit(1)

    # Check 3: Users have role assigned
    print("\n✓ CHECKING USER ROLES...")
    cursor.execute('SELECT username, role FROM users LIMIT 5;')
    users = cursor.fetchall()
    if not users:
        print("  ✗ ERROR: No users found in database!")
        sys.exit(1)
    
    for user in users:
        role = user.get('role', 'NOT SET')
        status = "✓" if role and role != 'NOT SET' else "✗"
        print(f"  {status} {user['username']}: role={role}")

    # Check 4: Count records
    print("\n✓ RECORD COUNTS...")
    cursor.execute('SELECT COUNT(*) as count FROM users;')
    user_count = cursor.fetchone()['count']
    print(f"  ✓ Total users: {user_count}")

    cursor.execute('SELECT COUNT(*) as count FROM password_resets;')
    reset_count = cursor.fetchone()['count']
    print(f"  ✓ Password resets recorded: {reset_count}")

    conn.close()

    # Final verdict
    print("\n" + "="*50)
    print("✅ STEP 1 VERIFICATION: SUCCESS!")
    print("="*50)
    print("\nYour database is ready. Proceed to STEP 2:")
    print("  → Create .env file with Gmail credentials")
    print("\n")

except mysql.connector.Error as err:
    print(f"\n✗ DATABASE ERROR: {err}")
    sys.exit(1)
except Exception as e:
    print(f"\n✗ UNEXPECTED ERROR: {e}")
    sys.exit(1)
