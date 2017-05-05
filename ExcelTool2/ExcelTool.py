#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys

import traceback

from src.Book import *
import time


inBook = None
outBook = None

def do():
    print("请等待...")

    global inBook
    global outBook

    inBook=None
    outBook=None



    ret, err, inBook = loadBook("in.xlsx")
    if ret != 0:
        return ret, err
    
    if not inBook.hasSheet("Sheet1"):
        return 1, "in.xlsx中不存在Sheet1"
    if not inBook.hasSheet("Sheet2"):
        return 1, "in.xlsx中不存在Sheet2"

    ret,err,inSheet1 = inBook.sheet("Sheet1")
    if ret != 0:
        return ret, err

    print("初始化Sheet1数字列序号")
    ret , err = inSheet1.initNumColIndex()
    if ret != 0:
        return ret, err

    
    ret,err,inSheet2 = inBook.sheet("Sheet2")
    if ret != 0:
        return ret, err

    print("初始化Sheet2数字列序号")
    ret,err = inSheet2.initNumColIndex()
    if ret != 0:
        return ret, err
    
    outBook = createBook()
    
    ret,err,outSheet = outBook.active()
    if ret != 0:
        return ret, err


    sheet1RowDone = {}
    sheet2RowDone = {}

    sheet1MaxRow = inSheet1.maxRow
    sheet2MaxRow = inSheet2.maxRow
    sheet1MaxCol = inSheet1.maxCol
    sheet2MaxCol = inSheet2.maxCol
    
    for rowIndex in range(1, sheet1MaxRow + 1):
        sheet1RowDone[rowIndex] = False

    for rowIndex in range(1, sheet2MaxRow + 1):
        sheet2RowDone[rowIndex] = False

    outSheetCount = 0

    print("初始化Sheet1数字列数据")
    ret,err = inSheet1.initNumColDict()
    if ret != 0:
        return ret, err

    print("初始化Sheet2数字列数据")
    ret,err = inSheet2.initNumColDict()
    if ret != 0:
        return ret, err

    print("初始化Sheet1唯一数字列")
    ret,err =inSheet1.initDiffNumRows()
    if ret != 0:
        return ret, err

    print("初始化Sheet2唯一数字列")
    ret,err =inSheet2.initDiffNumRows()
    if ret != 0:
        return ret, err

    print("正在处理唯一不同的树值")
    for sheet1RowIndex in inSheet1.diffNumRows:

        if sheet1RowDone[sheet1RowIndex] == True:
            continue

        for sheet2RowIndex in inSheet2.diffNumRows:
            if sheet2RowDone[sheet2RowIndex] == True:
                continue
            
            

            sheet1Num = inSheet1.numRowDict[sheet1RowIndex]
            sheet2Num = inSheet2.numRowDict[sheet2RowIndex]
            

            if sheet1Num == sheet2Num:
                sheet1RowDone[sheet1RowIndex] = True
                sheet2RowDone[sheet2RowIndex] = True
                
                outSheetCount +=1
                for sheet1ColIndex in range(1,sheet1MaxCol+1):
                    sheet1Cell = inSheet1.cell(sheet1RowIndex,\
                                         sheet1ColIndex)

                    outCell = outSheet.cell(outSheetCount,\
                                         sheet1ColIndex)
                    outCell.setVal(sheet1Cell.getVal())

                    outCell.setBlue()
                    
                    outCell.setTextFormat()
                    

                outSheetCount +=1
                for sheet2ColIndex in range(1,sheet2MaxCol+1):
                    sheet2Cell = inSheet2.cell(sheet2RowIndex,\
                                         sheet2ColIndex)
                    outCell = outSheet.cell(outSheetCount,\
                                         sheet2ColIndex)
                    outCell.setVal(sheet2Cell.getVal())

                    outCell.setRed()
                    outCell.setTextFormat()

    vals = []

    

    for sheet1Val in inSheet1.numValSet:
        sheet1Count = inSheet1.numValDict[sheet1Val]
        if sheet1Count ==1:
            continue
        for sheet2Val in inSheet2.numValSet:
            sheet2Count = inSheet2.numValDict[sheet2Val]
            if sheet2Count ==1:
                continue

            if sheet1Val != sheet2Val:
                continue

            if sheet1Count != sheet2Count:
                continue

            vals.append(sheet1Val)


    print("正在处理相同的树值")
    for val in vals:


        for sheet1RowIndex in inSheet1.getRowListByVal(val):
            if sheet1RowDone[sheet1RowIndex] == True:
                
                continue

            sheet1RowDone[sheet1RowIndex] = True
                   
            outSheetCount +=1
    

            for sheet1ColIndex in range(1,sheet1MaxCol+1):
                sheet1Cell = inSheet1.cell(sheet1RowIndex,\
                                             sheet1ColIndex)
                outCell = outSheet.cell(outSheetCount,\
                                             sheet1ColIndex)
                outCell.setVal(sheet1Cell.getVal())

                outCell.setBlue()
                outCell.setTextFormat()

        for sheet2RowIndex in inSheet2.getRowListByVal(val):
            if sheet2RowDone[sheet2RowIndex] == True:
                
                continue
            outSheetCount +=1
           
            sheet2RowDone[sheet2RowIndex] = True
            for sheet2ColIndex in range(1,sheet2MaxCol+1):
                sheet2Cell = inSheet2.cell(sheet2RowIndex,\
                                             sheet2ColIndex)
                outCell = outSheet.cell(outSheetCount,\
                                             sheet2ColIndex)
                outCell.setVal(sheet2Cell.getVal())

                outCell.setRed()
                outCell.setTextFormat()
                    
                
                

    
    print("正在处理没有匹配的项目")
    for sheet1RowIndex in range(1,sheet1MaxRow+1):

        if sheet1RowDone[sheet1RowIndex] == True:
            
            continue
        
        outSheetCount +=1

        sheet1RowDone[sheet1RowIndex] = True
        for sheet1ColIndex in range(1,inSheet1.maxCol+1):
                sheet1Cell = inSheet1.cell(sheet1RowIndex,\
                                         sheet1ColIndex)
                outCell = outSheet.cell(outSheetCount,\
                                         sheet1ColIndex)
                outCell.setVal(sheet1Cell.getVal())
                
                outCell.setBlue()

                outCell.setTextFormat()
                


    for sheet2RowIndex in range(1,sheet2MaxRow+1):

        if sheet2RowDone[sheet2RowIndex] == True:
            
            continue
                  
        sheet2RowDone[sheet2RowIndex] = True
        outSheetCount +=1
        for sheet2ColIndex in range(1,inSheet2.maxCol+1):
                sheet2Cell = inSheet2.cell(sheet2RowIndex,\
                                         sheet2ColIndex)
                outCell = outSheet.cell(outSheetCount,\
                                         sheet2ColIndex)
                outCell.setVal(sheet2Cell.getVal())
                

                outCell.setRed()

                outCell.setTextFormat()
    
    
    outBook.save("out.xlsx")

    return 0,None


            
        
        

    
        
    
if __name__=="__main__":
    try:
        ret, err = do()
        if ret != 0:
            print(err)
            if inBook is not None:
                inBook.close()
            if outBook is not None:
            
                outBook.close()
            time.sleep(2)
            
            sys.exit(1) 

        inBook.close()
        outBook.close()
        print("success")
        time.sleep(2)
    except Exception:
        print(traceback.format_exc())
        
        time.sleep(20)
    finally:
        if inBook is not None:
            inBook.close()
        if outBook is not None:
            
            outBook.close()

