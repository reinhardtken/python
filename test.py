import os
import sqlite3
import time
import util.excel
import util.my_sqlite






if __name__ == "__main__":
  """
  a = (3,)
  a2 = a + (3.4, 5)
  print(a[-1])
  print(a2[-1])
  print(time.mktime.__doc__)
  print(int(123.4))
  """

  
  all_start_time = time.time()
  xls = util.excel.EasyExcel(r'L:\workspace\github\mine\test\raw data\IF 5分钟 - 副本.xlsx')
  xls.SetActiveSheet("Sheet2")
  re = xls.GetRows4(10000, 4)
  cost_time = time.time() - all_start_time
  print(cost_time, len(re))
  if len(re) < 100:
    print(re)

  start_time = time.time()
  util.my_sqlite.WriteIF5ToDb(r"L:\workspace\github\mine\test\raw data\if5.db", re)
  cost_time = time.time() - start_time
  print(cost_time)
  print("all time", time.time() - all_start_time)
