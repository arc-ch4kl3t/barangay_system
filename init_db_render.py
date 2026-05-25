#!/usr/bin/env python3
"""
Initialize PostgreSQL database for Render deployment.
Run this ONCE after deploying to Render to create all necessary tables.

Usage:
    python init_db_render.py
"""

import os
import sys
import psycopg2
from psycopg2.extras import RealDictCursor

def init_database():
    """Initialize database schema with all required tables."""
    
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        print("ERROR: DATABASE_URL environment variable not set!")
        sys.exit(1)
    
    try:
        print("Connecting to PostgreSQL database...")
        conn = psycopg2.connect(database_url, sslmode="require")
        cursor = conn.cursor(cursor_factory=RealDictCursor)
        print("✓ Connected successfully")
        
        # Create users table
        print("\nCreating users table...")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(255) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                email VARCHAR(255),
                role VARCHAR(20) DEFAULT 'user',
                status VARCHAR(20) DEFAULT 'approved',
                signup_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        print("✓ Users table created")
        
        # Create password_resets table
        print("Creating password_resets table...")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS password_resets (
                id SERIAL PRIMARY KEY,
                username VARCHAR(100) NOT NULL,
                token VARCHAR(255) UNIQUE NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP NOT NULL,
                used BOOLEAN DEFAULT FALSE,
                used_at TIMESTAMP NULL,
                ip_address VARCHAR(80),
                attempt_count INT DEFAULT 1
            )
        """)
        print("✓ Password resets table created")
        
        # Create indexes on password_resets
        print("Creating indexes...")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_token ON password_resets (token)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_username ON password_resets (username)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_used ON password_resets (used)")
        print("✓ Indexes created")
        
        # Create households table
        print("Creating households table...")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS households (
                id SERIAL PRIMARY KEY,
                household_head VARCHAR(255),
                address VARCHAR(500),
                barangay VARCHAR(100),
                city VARCHAR(100),
                status VARCHAR(20) DEFAULT 'active',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        print("✓ Households table created")
        
        # Create household_members table
        print("Creating household_members table...")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS household_members (
                id SERIAL PRIMARY KEY,
                household_id INT NOT NULL REFERENCES households(id),
                full_name VARCHAR(255) NOT NULL,
                age INT,
                gender VARCHAR(20),
                relationship VARCHAR(50),
                occupation VARCHAR(100),
                status VARCHAR(20) DEFAULT 'active',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        print("✓ Household members table created")
        
        # Create audit_logs table
        print("Creating audit_logs table...")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS audit_logs (
                id SERIAL PRIMARY KEY,
                user_id VARCHAR(100),
                username VARCHAR(100) NOT NULL,
                action_type VARCHAR(50) NOT NULL,
                target_type VARCHAR(100) DEFAULT 'System',
                target_id VARCHAR(100) DEFAULT 'N/A',
                old_value TEXT,
                new_value TEXT,
                details TEXT,
                household_context VARCHAR(100) DEFAULT 'N/A',
                status VARCHAR(30) DEFAULT 'SUCCESS',
                ip_address VARCHAR(80),
                user_agent VARCHAR(255),
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        print("✓ Audit logs table created")
        
        # Create an index on audit_logs for better query performance
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_username_audit ON audit_logs (username)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_action_audit ON audit_logs (action_type)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_timestamp_audit ON audit_logs (timestamp)")
        print("✓ Audit log indexes created")
        
        # Check if admin user exists
        cursor.execute("SELECT COUNT(*) as count FROM users WHERE role='admin'")
        admin_count = cursor.fetchone()['count']
        
        if admin_count == 0:
            print("\nNo admin user found. Creating default admin user...")
            cursor.execute("""
                INSERT INTO users (username, password, email, role, status)
                VALUES (%s, %s, %s, %s, %s)
            """, ('admin', 'admin123', 'admin@barangay.local', 'admin', 'approved'))
            print("✓ Default admin user created")
            print("  USERNAME: admin")
            print("  PASSWORD: admin123")
            print("  ⚠️  IMPORTANT: Change this password immediately after first login!")
        else:
            print(f"\n✓ Admin user already exists ({admin_count} admin(s) found)")
        
        conn.commit()
        conn.close()
        
        print("\n" + "="*50)
        print("✓ DATABASE INITIALIZATION COMPLETE!")
        print("="*50)
        print("\nYour database is ready. You can now:")
        print("1. Login with admin/admin123 (if created)")
        print("2. Create user accounts")
        print("3. Manage households and residents")
        
        return True
        
    except psycopg2.errors.UndefinedTable as e:
        print(f"✗ Table error: {e}")
        sys.exit(1)
    except psycopg2.Error as e:
        print(f"✗ Database error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"✗ Unexpected error: {e}")
        sys.exit(1)

if __name__ == '__main__':
    init_database()
