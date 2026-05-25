import sys
import os
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    
    database_url = os.environ.get("DATABASE_URL")
    if not database_url:
        print("❌ DATABASE_URL environment variable not set!")
        sys.exit(1)
    
    conn = psycopg2.connect(database_url, sslmode="require")
    cursor = conn.cursor(cursor_factory=RealDictCursor)
    cursor.execute('SELECT id, username, role, email FROM users LIMIT 20;')
    results = cursor.fetchall()
    print('id\tusername\trole\temail')
    for row in results:
        print(f'{row["id"]}\t{row["username"]}\t{row["role"]}\t{row["email"]}')
    cursor.close()
    conn.close()
except ImportError:
    print('psycopg2 not available')
except Exception as e:
    print(f'Error: {e}')
