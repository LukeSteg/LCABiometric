#!/usr/bin/python
import datetime
import xlrd
import json

DATA_SHEET_NUM = 25
AGGREGATE_SHEET_NAME = "Aggregate"

def most_recent_key(tup):
    return str(tup[1]) + str(tup[0])

class genericFactory(object):

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = shapeRef.top
        self.y = shapeRef.left
        self.cx = shapeRef.width
        self.cy = shapeRef.height
        self.relBook = 0
        self.colNum = 0
        
    def getFileFromDict(self, fileDict):
        sortedFiles = sorted(fileDict, key = lambda x: most_recent_key(fileDict[x]), reverse = True)
        return sortedFiles[self.relBook]      

    def generateShape(self):
        print "WARNING generic generate shape invoked" 

    def getAggregateSheetFromBook(self, book):
        sheetName
   
    def setBook(self, bookNumber):
        self.relBook = bookNumber 

#    def setColumn(self,colNumber):
#        self.columnNum = colNumber
        
    def setColumn(self, columnName, fileDict):
        fileRef = self.getFileFromDict(fileDict)
        book = xlrd.open_workbook(fileRef)
        sheet = book.sheet_by_name(AGGREGATE_SHEET_NAME)
        columnNumber = 0
        readColumnText = sheet.cell_value(rowx = 0, colx = columnNumber)
        while((readColumnText.strip().upper() != columnName) and (readColumnText != '')):
            print readColumnText
            columnNumber += 1
            readColumnText = sheet.cell_value(rowx = 0, colx = columnNumber)

        if readColumnText == '':
            print 'WARNING, column title: ', readColumnText,  ' not found'

        print 'colnum ',columnNumber ,' readColumnText ', readColumnText
        
        self.colNum = columnNumber

    def setShape(self, shapeRef):
        self.shapeRef = shapeRef

    def setX(self, x):
        self.x = x 

    def setY(self, y):
        self.y = y 

    def setCX(self, cx):
        self.cx = cx 

    def setCY(self, CY):
        self.CY = CY

                                                                                            
