#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import openpyxl


class SrcCell():
    def __init__(self,val):
        self.val=val
        pass
    def getVal(self):
        return self.val

class SrcRow():
    def __init__(self):
        pass
    def getNumCell(self):
        pass
    def getVal(self):
        pass
    def EqualVal(self):
        pass

class SrcSheet():
    def __init__(self,sheet):
        self.rows=[]
        maxRow = sheet.max_row
        for index in range(1,maxRow+1):
            row = SrcRow()
            self.rows.append(SrcRow)
        self.rows = rows
        pass
    def getRow(self):
        pass
    def countNumCell(self):
        pass
    def setRowDone(self):
        pass
    def sort(self):
        pass
    def getDiffRows(self):
        pass
    def getFreeRows(self):
        pass
    def getDupliVals(self):
        pass
    def getRowsByVal(self):
        pass
    def getRowsInTotal(self):
        pass
    

class SrcBook():
    def __init__(self,path):
        if not os.path.isfile(path):
            raise MyError("源文件 %s 不存在" % path)
        book = openpyxl.load_workbook(path,read_only=True, \
                                    keep_vba=False, data_only=True, \
                                    guess_types=False, keep_links=False)
        
        self.sheet1 =SrcSheet(book["Sheet1"])
        self.sheet2 =SrcSheet(book["Sheet2"])
        pass
    def getSheetByName(self,sheetName):
        pass
    def addSheet(self):
        pass
    def findEqual(self):
        pass
    def compSheet(self):
        pass

class TarCell():
    def __init__(self,color):
        self.color = color
        pass
    def setColor(self,color):
        self.color=color
        pass

class TarRow():
    def __init__(self,cells):
        self.cells=cells
        pass
    def setColor(self,color):
        for cell in cells:
            cell.setColor(color)
        pass

class TarSheet():
    def __init__(self,rows):
        self.rows=rows
        pass
    def addRow(self,row):
        self.rows.append(row)
        row.setColor(color)
        pass

class TarBook():
    def __init__(self):

        pass
    def getSheet(self):
        pass
    def write(self,path):
        pass

def run():
    srcBook = SrcBook("in.xlsx")
    srcSheet1 = srcBook.getSheet()
    srcSheet2 = srcBook.getSheet()

    tarBook = TarBook()
    tarSheet = tarBook.getSheet()

    for row1 in sheet1.getDiffRows():
        for row2 in sheet2.getDiffRows():
            if row1.equalVal(row2):
                tarSheet.addRow(row1,color1)
                tarSheet.addRow(row2,color2)
                row1.done()
                row2.done()
    for val in sheet1.getDupliVals():

        if sheet1.countNumCell(val)==sheet2.countNumCell(val):
            for row1 in sheet1.getRowsByVal(row1Val):
                tarSheet.addRow(row1,color1)
                row1.done()
            for row2 in sheet1.getRowsByVal(row1Val):
                tarSheet.addRow(row2,color2)
                row2.done()

    for row1 in sheet1.getDiffRows():
        rows = sheet2.getRowsInTotal(row1.getVal())
        if len(rows) == 0:
            continue
        tarSheet.addRow(row1,color1)
        row1.done()
        for row in rows:
            tarSheet.addRow(row2,color2)
            row2.done()

    for row2 in sheet2.getDiffRows():
        rows = sheet1.getRowsInTotal(row2.getVal())
        if len(rows) == 0:
            continue
        tarSheet.addRow(row2,color2)
        row2.done()
        for row in rows:
            tarSheet.addRow(row1,color1)
            row1.done()
                

    for row1 in sheet1.getFreeRows():
        tarSheet.addRow(row1,color1)
        row1.done()

    for row2 in sheet2.getFreeRows():
        tarSheet.addRow(row2,color2)
        row2.done()


    tarBook.write(tarPath)
    

class ExcelStat():
    def __init__(self,index,val):
        self.index=index
        self.val=val
        self.done=False

class MyError(Exception):
    pass

class ExcelService(object):
    def __init__(self,srcPath="in.xlsx",tarPath="out.xlsx"):
        if not os.path.isfile(srcPath):
            raise MyError("源文件 %s 不存在" % srcPath)
        if os.path.isfile(tarPath):
            raise MyError("请删除目标文件 %s" % tarPath)
        self.srcPath = srcPath
        self.tarPath = tarPath

        self.srcBook = openpyxl.load_workbook(srcPath,read_only=True, \
                                    keep_vba=False, data_only=True, \
                                    guess_types=False, keep_links=False)
        self.tarBook = openpyxl.Workbook(write_only=True)
        

    def saveTarBook(self):
        self.tarBook.save(self.tarPath)

    def close(self):
        self.srcBook.close()
        self.tarBook.close()

        self.srcBook = None
        self.tarBook = None

        
        

def main():
    service = ExcelService()
    service.close()
    
if __name__=="__main__":
    run()

