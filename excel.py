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
              self.xlBook = self.xlApp.Workbooks.Open(filename)
          else:
              self.xlBook = self.xlApp.Workbooks.Add()
              self.filename = ''

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

      def GetRange(self, sheet, row1, col1, row2, col2):
          "return a 2d array (i.e. tuple of tuples)"
          sht = self.xlBook.Worksheets(sheet)
          #print(sht.__dict__)
          return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value

      def AddPicture(self, sheet, pictureName, Left, Top, Width, Height):
          "Insert a picture in sheet"
          sht = self.xlBook.Worksheets(sheet)
          sht.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)

      def CopySheet(self, before):
          "copy sheet"
          shts = self.xlBook.Worksheets
          shts(1).Copy(None,shts(1))

      def UsedRowCount(self, sheet):
        sht = self.xlBook.Worksheets(sheet)
        return sht.UsedRange.Rows.Count

      def UsedColCount(self, sheet):
        sht = self.xlBook.Worksheets(sheet)
        return sht.UsedRange.Columns.Count

      def GetUsedSheetValue(self, sheet):
        return self.GetRange(sheet, 1, 1, self.UsedRowCount(sheet), self.UsedColCount(sheet))



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
      xls = EasyExcel(r'N:\workspace\github\temp\test - 副本.xlsx')
      i = 4
      used_row_num = xls.UsedRowCount("Sheet1")
      used_col_num = xls.UsedColCount("Sheet1")

      re = []
      while i <= used_row_num:
        a = (xls.GetRange("Sheet1", i , 1, i, used_col_num)[0],)
        print(a)
        b = a + time.mktime(time.strptime(a[0],'%Y-%m-%d %H:%M:%S'))
        re.append(b)
        #print(xls.GetRange("Sheet1", i , 1, i, used_col_num))
        i += 1
      print(re)


      #re = time.mktime(time.strptime(b,'%Y-%m-%d %H:%M:%S'))
      #print(xls.__dict__)
      #xls.Save()
      xls.Close()
