#!/usr/bin/env python3
"""
Apply self-registration system migration to the database
Adds status and signup_date columns to users table
"""
import mysql.connector

def run_migration():
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="barangay_db"
    )
    cursor = conn.cursor()
    
    print("🔄 Applying Self-Registration Migration...")
    print("-" * 60)
    
    try:
        # 1. Add status column
        print("✓ Adding 'status' column to users table...")
        cursor.execute("""
            ALTER TABLE users ADD COLUMN status VARCHAR(20) DEFAULT 'approved' 
            COMMENT 'pending, approved, or rejected'
        """)
        
        # 2. Add signup_date column
        print("✓ Adding 'signup_date' column to users table...")
        cursor.execute("""
            ALTER TABLE users ADD COLUMN signup_date TIMESTAMP NULL 
            COMMENT 'When user self-registered'
        """)
        
        # 3. Update existing users to approved status
        print("✓ Setting existing users to 'approved' status...")
        cursor.execute("""
            UPDATE users SET status = 'approved' 
            WHERE status IS NULL OR status = ''
        """)
        
        # 4. Create indexes
        print("✓ Creating database indexes...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_status ON users(status)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_signup_date ON users(signup_date)")
        
        conn.commit()
        
        # Verify
        print("\n" + "=" * 60)
        print("✅ MIGRATION COMPLETE!")
        print("=" * 60)
        
        cursor.execute("SHOW COLUMNS FROM users")
        columns = cursor.fetchall()
        print("\n📋 Users table columns:")
        for col in columns:
            print(f"   • {col[0]}")
        
        cursor.execute("SELECT COUNT(*) as total FROM users WHERE status='approved'")
        count = cursor.fetchone()[0]
        print(f"\n✓ {count} users set to 'approved' status")
        
    except mysql.connector.Error as e:
        if "1060" in str(e):
            print("⚠️  Columns already exist - skipping...")
            print("   (Migration may have already been applied)")
        else:
            print(f"❌ Error: {e}")
            return False
    
    finally:
        cursor.close()
        conn.close()
    
    return True

if __name__ == "__main__":
    success = run_migration()
    if success:
        print("\n🎉 Your system now supports self-registration!")
        print("\n📝 Next steps:")
        print("   1. Restart Flask: python app.py")
        print("   2. Visit http://localhost:5000 to test")
        print("   3. Users can now sign up at /signup")
        print("   4. Check /user-management to approve/reject signups")
