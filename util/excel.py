#!/usr/bin/env python
# -*- coding: utf-8 -*-
from win32com.client import Dispatch
import win32com.client
import time

class EasyExcel:
      """A utility to make it easier to get at Excel.    Remembering
      to save the data is your problem, as is    error handling.
      Operates on one workbook at a time."""

      def __init__(self, filename=None):
          self.xlApp = win32com.client.DispatchEx('Excel.Application')
          if filename:
              self.filename = filename
              #print(self.xlApp.__dict__)
              start_time = time.time()
              self.xlBook = self.xlApp.Workbooks.Open(filename)
              print("self.xlApp.Workbooks.Open time", time.time() - start_time)
          else:
              self.xlBook = self.xlApp.Workbooks.Add()
              self.filename = ''

          self.SetActiveSheet(None)

      def Save(self, newfilename=None):
          if newfilename:
              self.filename = newfilename
              self.xlBook.SaveAs(newfilename)
          else:
              self.xlBook.Save()

      def Close(self):
          self.xlBook.Close(SaveChanges=0)
          del self.xlApp

      def GetCell(self, sheet, row, col):
          "Get value of one cell"
          sht = self.xlBook.Worksheets(sheet)
          return sht.Cells(row, col).Value

      def SetCell(self, sheet, row, col, value):
          "set value of one cell"
          sht = self.xlBook.Worksheets(sheet)
          sht.Cells(row, col).Value = value


      def AddPicture(self, sheet, pictureName, Left, Top, Width, Height):
          "Insert a picture in sheet"
          sht = self.xlBook.Worksheets(sheet)
          sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

      def CopySheet(self, before):
          "copy sheet"
          shts = self.xlBook.Worksheets
          shts(1).Copy(None,shts(1))

      def UsedRowCount(self, sheet=None):
        if sheet:
          sht = self.xlBook.Worksheets(sheet)
          return sht.UsedRange.Rows.Count
        elif self.active_sheet:
          return self.active_sheet.UsedRange.Rows.Count
        else:
          print("error UsedRowCount")

      def UsedColCount(self, sheet=None):
        if sheet:
          sht = self.xlBook.Worksheets(sheet)
          return sht.UsedRange.Columns.Count
        elif self.active_sheet:
          return self.active_sheet.UsedRange.Columns.Count
        else:
          print("error UsedColCount")

      """
      def GetUsedSheetValue(self, sheet):
        return self.GetRange(sheet, 1, 1, self.UsedRowCount(sheet), self.UsedColCount(sheet))
      """

      def GetRange(self, row1, col1, row2, col2, sheet=None):
          #print("return a 2d array (i.e. tuple of tuples)")
          if sheet:
            #print("sheet")
            sht = self.xlBook.Worksheets(sheet)
            return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value
          elif self.active_sheet:
            #print("self sheet")
            return self.active_sheet.Range(self.active_sheet.Cells(row1, col1), self.active_sheet.Cells(row2, col2)).Value
          else:
            print("error UsedColCount")

      def SetActiveSheet(self, sheet):
        if sheet:
          self.active_sheet = self.xlBook.Worksheets(sheet)
        else:
          self.active_sheet = None


      def GetRows(self, start=0, sheet=None):
        re = []
        i = start
        if sheet:
          used_row_num = self.UsedRowCount(sheet)
          used_col_num = self.UsedColCount(sheet)
          print(used_row_num, used_col_num)
          while i < used_row_num:
            a = self.GetRange(i , 1, i, used_col_num, sheet)[0]
            re.append(a)
            #print(self.GetRange("Sheet1", i , 1, i, used_col_num))
            i += 1
        elif self.active_sheet:
          used_row_num = self.UsedRowCount()
          used_col_num = self.UsedColCount()
          print(used_row_num, used_col_num)
          while i < used_row_num:
            a = self.GetRange(i , 1, i, used_col_num)[0]
            re.append(a)
            #print(self.GetRange("Sheet1", i , 1, i, used_col_num))
            i += 1

        return re

      def GetRows2(self, start=0, sheet=None):
        "对返回数据做加工"
        re = []
        i = start
        if sheet:
          used_row_num = self.UsedRowCount(sheet)
          used_col_num = self.UsedColCount(sheet)
          print(used_row_num, used_col_num)
          while i < used_row_num:
            a = self.GetRange(i , 1, i, used_col_num, sheet)[0]
            #print(a)
            b = a + (int(time.mktime(time.strptime(a[0],'%Y-%m-%d %H:%M:%S'))),)
            #print(b)
            re.append(b)
            #print(self.GetRange("Sheet1", i , 1, i, used_col_num))
            i += 1
        elif self.active_sheet:
          used_row_num = self.UsedRowCount()
          used_col_num = self.UsedColCount()
          print(used_row_num, used_col_num)
          while i < used_row_num:
            a = self.GetRange(i , 1, i, used_col_num)[0]
            #print(a)
            b = a + (int(time.mktime(time.strptime(a[0],'%Y-%m-%d %H:%M:%S'))),)
            #print(b)
            re.append(b)
            #print(self.GetRange("Sheet1", i , 1, i, used_col_num))
            i += 1

        return re

      def GetRows3(self, N=1000,start=0, sheet=None):
        "一次取n个数据"
        re = []
        i = start
        if sheet:
          used_row_num = self.UsedRowCount(sheet)
          used_col_num = self.UsedColCount(sheet)
          print(used_row_num, used_col_num)
          while i <= used_row_num:
            des = 0
            if i + N <= used_row_num:
              des = i + N
            else:
              des = used_row_num
            a = self.GetRange(i , 1, des, used_col_num, sheet)
            #print("GetRows3.a", a)
            for v in a:
              re.append(v)
            #print(self.GetRange("Sheet1", i , 1, i, used_col_num))
            print("GetRows3.i des N", i, des, N)
            i = des+1

        elif self.active_sheet:
          used_row_num = self.UsedRowCount()
          used_col_num = self.UsedColCount()
          print(used_row_num, used_col_num)
          while i <= used_row_num:
            des = 0
            if i + N <= used_row_num:
              des = i + N
            else:
              des = used_row_num
            start_time = time.time()
            a = self.GetRange(i , 1, des, used_col_num)
            print("self.GetRange time", time.time() - start_time)
            #print("GetRows3.a", a, "\r\n")
            start_time = time.time()
            for v in a:
              re.append(v)
            print("for v in a time", time.time() - start_time)
            #print(self.GetRange("Sheet1", i , 1, i, used_col_num))
            #print("GetRows3.i des N", i, des, N)

            i = des+1
            

        return re


      def GetRows4(self, N=1000,start=0, sheet=None):
        "一次取n个数据,并在最后编辑结果"
        re = self.GetRows3(N, start, sheet)
        print("GetRows4.len(re)", len(re))
        #print("GetRows4.re", re)
        re2 = []    
        start_time = time.time()
        for i in re:
          b = i + (int(time.mktime(time.strptime(i[0],'%Y-%m-%d %H:%M:%S'))),)
          re2.append(b)
        print("for i in re: time", time.time() - start_time)
        return re2

#==============================================================================

if __name__ == "__main__":
      """
      #PNFILE = r'c:\screenshot.bmp'
      xls = EasyExcel(r'N:\workspace\github\temp\IF 5m - 副本.xls')
      #xls.addPicture('Sheet1', PNFILE, 20,20,1000,1000)
      #xls.cpSheet('Sheet1')
      #print(xls.GetSheet("Sheet1"))
      #print(xls.GetRange("Sheet1", 1, 1, 5, 5))
      #print(xls.GetUsedSheetValue("Sheet2"))
      i = 1
      used_row_num = xls.UsedRowCount("Sheet2")
      used_col_num = xls.UsedColCount("Sheet2")
      while i <= used_row_num:
        #nonlocal i
        print(xls.GetRange("Sheet2", i , 1, i, used_col_num))
        i += 1
      """

      #PNFILE = r'c:\screenshot.bmp'
      #xls = EasyExcel(r'L:\workspace\github\mine\test\raw data\test.xlsx')
      """
      i = 4
      xls.SetActiveSheet("Sheet1")
      used_row_num = xls.UsedRowCount()
      used_col_num = xls.UsedColCount()
      print(used_row_num, used_col_num)
      re = []
      while i <= used_row_num:
        a = xls.GetRange(i , 1, i, used_col_num)[0]
        #print(a)
        b = a + (int(time.mktime(time.strptime(a[0],'%Y-%m-%d %H:%M:%S'))),)
        #print(b)
        re.append(b)
        #print(xls.GetRange("Sheet1", i , 1, i, used_col_num))
        i += 1
      print(re)
      """
      start_time = time.time()
      xls = EasyExcel(r'L:\workspace\github\mine\test\raw data\IF 5分钟.xlsx')
      xls.SetActiveSheet("Sheet2")
      re = xls.GetRows4(1, 4)
      cost_time = time.time() - start_time
      print(cost_time)
      print(re[:7])


      #re = time.mktime(time.strptime(b,'%Y-%m-%d %H:%M:%S'))
      #print(xls.__dict__)
      #xls.Save()
      xls.Close()
