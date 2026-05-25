#!/usr/bin/env python3
"""
Add email column to users table for Gmail integration
"""
import mysql.connector

def add_email_column():
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="barangay_db"
    )
    cursor = conn.cursor()
    
    print("🔄 Adding Email Column for Gmail Integration...")
    print("-" * 60)
    
    try:
        # Check if email column exists
        cursor.execute("SHOW COLUMNS FROM users LIKE 'email'")
        if cursor.fetchone():
            print("⚠️  Email column already exists")
            return True
        
        # Add email column
        print("✓ Adding 'email' column to users table...")
        cursor.execute("""
            ALTER TABLE users ADD COLUMN email VARCHAR(100) UNIQUE 
            COMMENT 'Email address for Gmail integration and notifications'
        """)
        
        # Create index for faster lookups
        print("✓ Creating email index...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_email ON users(email)")
        
        conn.commit()
        
        print("\n" + "=" * 60)
        print("✅ EMAIL COLUMN ADDED!")
        print("=" * 60)
        
        cursor.execute("SHOW COLUMNS FROM users WHERE Field='email'")
        col_info = cursor.fetchone()
        print(f"\n✓ Email column type: {col_info[1]}")
        print("✓ Email is UNIQUE - prevents duplicate registrations")
        
    except mysql.connector.Error as e:
        print(f"❌ Error: {e}")
        return False
    
    finally:
        cursor.close()
        conn.close()
    
    return True

if __name__ == "__main__":
    success = add_email_column()
    if success:
        print("\n🎉 Email integration ready!")
        print("\n✅ Your system can now:")
        print("   • Store resident email addresses")
        print("   • Send password recovery emails via Gmail SMTP")
        print("   • Send notification emails for signups/approvals")
        print("   • Prevent duplicate email registrations")
