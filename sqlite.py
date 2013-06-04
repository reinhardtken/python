import os
import sqlite3
import time





ddl_file = r"N:\workspace\github\mine\python\src\create_db.ddl"
db_file = r"N:\workspace\github\temp\if5.db"



#=================================================================
if __name__ == "__main__":
	if !os.path.exists(db_fille):
		with sqlite3.connect(db_file) as conn:
			with open(ddl_file, 'rt') as f:
				schema = f.read()
			conn.executescript(schema)

  with sqlite3.connect(db_file) as conn:
		cursor = conn.cursor()
    # Insert the objects into the database
    cursor.executemany("insert into obj (data) values (?)", to_save)

a = "2011-09-28 10:00:00"
b = "2013-03-01 13:05:00"
c = "2013-03-01 13:10:00"
re = time.mktime(time.strptime(b,'%Y-%m-%d %H:%M:%S'))
re2 = time.mktime(time.strptime(c,'%Y-%m-%d %H:%M:%S'))
print(re)
print(re2)
print(time.time())
