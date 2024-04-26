import psycopg2 as pg

try:
    conn = pg.connect(
        host='localhost',
        database='football_club',
        port=5432,
        user='postgres',
        password="Roma23082003"
    )
    cursor = conn.cursor()
    print("Connection established")
except Exception as err:
    print("Something went wrong")
    print(err)

def fetch_data():
    cursor.execute('''SELECT * FROM privet''')
    data = cursor.fetchall()
    return data

def create_entry():
    cursor.execute('''INSERT INTO privet (id, name, age) 
                    VALUES (%s, %s, %s) RETURNING *''', (11, 'My Ex', 22))
    add_data = cursor.fetchone()
    conn.commit()
    return add_data

def delete_entry():
    cursor.execute('''DELETE FROM privet
                      WHERE id = %s''',('1'))
    conn.commit()
    return 'Data deleted successfuly'

def update_entry():
    cursor.execute('''UPDATE privet
                           SET name = %s, age = %s WHERE id = %s''',
                   ('My_other exe', 21, 2))
    conn.commit()
    return 'Update database'

details = fetch_data()
for row in details:
    print(row)

kkk = create_entry()
print(kkk)

details1 = fetch_data()
for row in details1:
    print(row)

data = delete_entry()
print(data)

details3 = fetch_data()
for row in details3:
    print(row)

data = update_entry()
print(data)

det = fetch_data()
for row in det:
    print(row)