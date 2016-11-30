#!/usr/bin/python
import datetime
import xlrd
import json

DATA_SHEET_NUM = 25
AGGREGATE_SHEET_NAME = "Aggregate"

def most_recent_key(tup):
    month_dict = {'Jan':1, 'Feb':2, 'Mar':3, 'April':4, 'May':5, 'June':6, 'July':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}
    return str(tup[1]) + str(10+tup[0])
    #tuple is (month, year) and we want most recent to be the highest number, because we reverse the sort.
    #add ten to month so that it is of fixed length

class genericFactory(object):

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = shapeRef.top
        self.y = shapeRef.left
        self.cx = shapeRef.width
        self.cy = shapeRef.height
        self.relBook = 0
        
    def getFileFromDict(self, fileDict):
        sortedFiles = sorted(fileDict, key = lambda x: most_recent_key(fileDict[x]), reverse = True)
        return sortedFiles[self.relBook]      

    def generateShape(self):
        print "generic generate shape invoked" 

    def getAggregateSheetFromBook(self, book):
        sheetName
   
    def setBook(self, bookNumber):
        self.relBook = bookNumber 

    def setColumn(self,colText):
        self.columnNum = colText
        
    def setColumnName(self, colName):
        self.columnName = colName

    def setShape(self, shapeRef):
        self.shapeRef =  shapeRef

    def setX(self, x):
        self.x = x 

    def setY(self, y):
        self.y = y 

    def setCX(self, cx):
        self.cx = cx 

    def setCY(self, CY):
        self.CY = CY

                                                                                            
