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
        
    def getFileFromDict(self, fileDict):
        sortedFiles = sorted(fileDict, key = lambda x: most_recent_key(fileDict[x]), reverse = True)
        tempDict = {}
        for f in fileDict:
            fileDict[f]
            most_recent_key(fileDict[f])
            tempDict[most_recent_key(fileDict[f])] = f    
        
        #sortedFiles = sorted(tempDict)
        print 'relbk ',self.relBook
        print 'stdf', sortedFiles
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

                                                                                            
