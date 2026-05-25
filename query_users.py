import sys
try:
    import mysql.connector
    conn = mysql.connector.connect(
        host='localhost',
        user='root',
        password='',
        database='barangay_db'
    )
    cursor = conn.cursor()
    cursor.execute('SELECT id, username, role, email FROM users LIMIT 20;')
    results = cursor.fetchall()
    print('id\tusername\trole\temail')
    for row in results:
        print(f'{row[0]}\t{row[1]}\t{row[2]}\t{row[3]}')
    cursor.close()
    conn.close()
except ImportError:
    print('mysql.connector not available')
except Exception as e:
    print(f'Error: {e}')
