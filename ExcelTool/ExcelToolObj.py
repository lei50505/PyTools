#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys
sys.path.append("../lib")

import traceback

from OpenpyxlUtils import *




def do():

    inBook=None
    outBook=None

    try:
        if not os.path.isfile("in.xlsx"):
            raise Exception("in.xlsx不存在")

        inBook = loadBook("in.xlsx")
        
        if not inBook.hasSheet("Sheet1"):
            raise Exception("in.xlsx中不存在Sheet1")
        if not inBook.hasSheet("Sheet2"):
            raise Exception("in.xlsx中不存在Sheet2")

        inSheet1 = inBook.sheet("Sheet1")
        
        inSheet2 = inBook.sheet("Sheet2")
        
        outBook = createBook()
        outSheet = outBook.active


        

        thinSide = Side(border_style="thin", color="000000")
        thinBorder = Border(top=thinSide, left=thinSide, \
                            right=thinSide, bottom=thinSide)


        sheet1RowDone = {}
        sheet2RowDone = {}
        
        for rowIndex in range(1,inSheet1.maxRow+1):
            sheet1RowDone[rowIndex]=False

        for rowIndex in range(1,inSheet2.maxRow+1):
            sheet2RowDone[rowIndex]=False

        outSheetCount = 0

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
                    for sheet1ColIndex in range(1,inSheet1.maxCol+1):
                        sheet1Cell = inSheet1.cell(sheet1RowIndex,\
                                             sheet1ColIndex)
                        outCell = outSheet.cell(outSheetCount,\
                                             sheet1ColIndex)
                        outCell.setVal(sheet1Cell.val)

                        outCell.setBlue()
                        

                    outSheetCount +=1
                    for sheet2ColIndex in range(1,inSheet2.maxCol+1):
                        sheet2Cell = inSheet2.cell(sheet2RowIndex,\
                                             sheet2ColIndex)
                        outCell = outSheet.cell(outSheetCount,\
                                             sheet2ColIndex)
                        outCell.setVal(sheet2Cell.val)

                        outCell.setRed()

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

        for val in vals:
                

            for sheet1RowIndex in inSheet1.getRowListByVal(val):
                if sheet1RowDone[sheet1RowIndex] == True:
                    
                    continue

                sheet1RowDone[sheet1RowIndex] = True
                       
                outSheetCount +=1
        

                for sheet1ColIndex in range(1,inSheet1.maxCol+1):
                    sheet1Cell = inSheet1.cell(sheet1RowIndex,\
                                                 sheet1ColIndex)
                    outCell = outSheet.cell(outSheetCount,\
                                                 sheet1ColIndex)
                    outCell.setVal(sheet1Cell.val)

                    outCell.setBlue()

            for sheet2RowIndex in inSheet2.getRowListByVal(val):
                if sheet2RowDone[sheet2RowIndex] == True:
                    
                    continue
                outSheetCount +=1

                sheet2RowDone[sheet2RowIndex] = True
                for sheet2ColIndex in range(1,inSheet2.maxCol+1):
                    sheet2Cell = inSheet2.cell(sheet2RowIndex,\
                                                 sheet2ColIndex)
                    outCell = outSheet.cell(outSheetCount,\
                                                 sheet2ColIndex)
                    outCell.setVal(sheet2Cell.val)

                    outCell.setRed()
                        
                    
                    

        
        
        for sheet1RowIndex in range(1,inSheet1.maxRow+1):
            if sheet1RowDone[sheet1RowIndex] == True:
                
                continue
            
            outSheetCount +=1

            sheet1RowDone[sheet1RowIndex] = True
            for sheet1ColIndex in range(1,inSheet1.maxCol+1):
                    sheet1Cell = inSheet1.cell(sheet1RowIndex,\
                                             sheet1ColIndex)
                    outCell = outSheet.cell(outSheetCount,\
                                             sheet1ColIndex)
                    outCell.setVal(sheet1Cell.val)

                    outCell.setBlue()
                    


        for sheet2RowIndex in range(1,inSheet2.maxRow+1):
            if sheet2RowDone[sheet2RowIndex] == True:
                continue
                        
            sheet2RowDone[sheet1RowIndex] = True
            outSheetCount +=1
            for sheet2ColIndex in range(1,inSheet2.maxCol+1):
                    sheet2Cell = inSheet2.cell(sheet2RowIndex,\
                                             sheet2ColIndex)
                    outCell = outSheet.cell(outSheetCount,\
                                             sheet2ColIndex)
                    outCell.setVal(sheet2Cell.val)

                    outCell.setRed()
        
        
        outBook.save("out.xlsx")
        inBook.close()
        outBook.close()
        print("success")
    except Exception:
        
        print(traceback.format_exc())
    finally:
        inBook.close()
        outBook.close()
        
        

    #time.sleep(2)
        
    
if __name__=="__main__":
    do()

