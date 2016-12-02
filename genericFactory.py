#!/usr/bin/python
import datetime
import xlrd
import json
from pptx.util import Inches

DATA_SHEET_NUM = 25
AGGREGATE_SHEET_NAME = "Aggregate"

def most_recent_key(tup):
    month_dict = {'Jan':1, 'Feb':2, 'Mar':3, 'April':4, 'May':5, 'June':6, 'July':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}
    return str(tup[1]) + str(10+month_dict[str(tup[0])])
    #tuple is (month, year) and we want most recent to be the highest number, because we reverse the sort.
    #add ten to month so that it is of fixed length

class genericFactory(object):

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = Inches(shapeRef.top)
        self.y = Inches(shapeRef.left)
        self.cx = Inches(shapeRef.width)
        self.cy = Inches(shapeRef.height)
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
            print readColumnText.strip().upper() + ' ; ' + str(columnName)
            columnNumber += 1
            readColumnText = sheet.cell_value(rowx = 0, colx = columnNumber)

        if readColumnText == '':
            print 'WARNING, column title: ', readColumnText,  ' not found'

        print 'colnum ',columnNumber ,' readColumnText ', readColumnText
        
        self.colNum = columnNumber

    def setShape(self, shapeRef):
        self.shapeRef = shapeRef

    def setX(self, x):
        self.x = Inches(float(x))

    def setY(self, y):
        self.y = Inches(float(y)) 

    def setCX(self, cx):
        self.cx = Inches(float(cx)) 

    def setCY(self, cy):
        self.cy = Inches(float(cy))

                                                                                            
