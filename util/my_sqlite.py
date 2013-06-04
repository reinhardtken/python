import os
import sqlite3
import time





ddl_file = r"L:\workspace\github\mine\python\src\create_db.ddl"
db_file = r"L:\workspace\github\mine\test\raw data\if5.db"


test_data = [('2010-05-04 09:20:00', 3087.0, 3075.2, 3079.8, 3083.4, 14.2, 45.34, 5.33, 1272936000), ('2010-05-04 09:45:00', 3083.6, 3076.2, 3082.4, 3076.6, 12.93, 44.7, -4.07, 1272937500), ('2010-05-04 10:35:00', 3092.6, 3081.6, 3081.6, 3090.0, 8.65, 56.52, 0.5, 1272940500), ('2010-05-04 13:45:00', 3077.8, 3073.0, 3076.4, 3074.0, 6.87, 43.74, -1.44, 1272951900), ('2010-05-05 11:05:00', 3040.4, 3037.0, 3037.6, 3038.6, 4.99, 39.16, 0.29, 1273028700), ('2010-05-06 14:45:00', 2975.0, 2966.2, 2974.0, 2972.0, 11.71, 21.86, -4.98, 1273128300), ('2010-05-11 14:35:00', 2819.4, 2803.0, 2819.4, 2815.0, 12.89, 20.24, -8.72, 1273559700)]


def WriteIF5ToDb(file, data):
  if ~os.path.exists(file):
    with sqlite3.connect(file) as conn:
      with open(ddl_file, 'rt') as f:
        schema = f.read()
      conn.executescript(schema)

  with sqlite3.connect(file) as conn:
    cursor = conn.cursor()
    # Insert the objects into the database
    cursor.executemany("insert into IF5(time_str, high_value, low_value,  start_value,  end_value,  dif_value,  rsi_value,  macd_value, time) \
      values (?,?,?,?,?,?,?,?,?)", data)

#=================================================================
if __name__ == "__main__":
  """
  if ~os.path.exists(db_file):
    with sqlite3.connect(db_file) as conn:
      with open(ddl_file, 'rt') as f:
        schema = f.read()
      conn.executescript(schema)

  with sqlite3.connect(db_file) as conn:
    cursor = conn.cursor()
    # Insert the objects into the database
    cursor.executemany("insert into IF5(time_str, high_value,	low_value,	start_value,	end_value,	dif_value,	rsi_value,	macd_value,	time) \
      values (?,?,?,?,?,?,?,?,?)", test_data)
  """
  WriteIF5ToDb(db_file, test_data)
  a = "2011-09-28 10:00:00"
  b = "2013-03-01 13:05:00"
  c = "2013-03-01 13:10:00"
  re = time.mktime(time.strptime(b,'%Y-%m-%d %H:%M:%S'))
  re2 = time.mktime(time.strptime(c,'%Y-%m-%d %H:%M:%S'))
  print(re)
  print(re2)
  print(time.time())
	 