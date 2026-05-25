#!/usr/bin/env python3
"""
Add email column to users table for Gmail integration
"""
import os
import psycopg2

def add_email_column():
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        print("❌ DATABASE_URL environment variable not set!")
        return False
    
    conn = psycopg2.connect(database_url, sslmode="require")
    cursor = conn.cursor()
    
    print("🔄 Adding Email Column for Gmail Integration...")
    print("-" * 60)
    
    try:
        # Check if email column exists
        cursor.execute("""
            SELECT column_name FROM information_schema.columns 
            WHERE table_name='users' AND column_name='email'
        """)
        if cursor.fetchone():
            print("⚠️  Email column already exists")
            return True
        
        # Add email column
        print("✓ Adding 'email' column to users table...")
        cursor.execute("""
            ALTER TABLE users ADD COLUMN email VARCHAR(100) UNIQUE
        """)
        
        # Create index for faster lookups
        print("✓ Creating email index...")
        try:
            cursor.execute("CREATE INDEX idx_email ON users(email)")
        except psycopg2.Error:
            pass
        
        conn.commit()
        
        print("\n" + "=" * 60)
        print("✅ EMAIL COLUMN ADDED!")
        print("=" * 60)
        
        cursor.execute("""
            SELECT data_type FROM information_schema.columns 
            WHERE table_name='users' AND column_name='email'
        """)
        col_info = cursor.fetchone()
        print(f"\n✓ Email column type: {col_info[0]}")
        print("✓ Email is UNIQUE - prevents duplicate registrations")
        
    except psycopg2.Error as e:
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
