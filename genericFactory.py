#!/usr/bin/python
import datetime
import xlrd

DATA_SHEET_NUM = 25

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
        
    def setBook(self, yr):
        self.relYear = yr
        
    def getFileFromDict(self, fileDict):
        sortedFiles = sorted(fileDict, key = lambda x: most_recent_key(fileDict[x]), reverse = True)
        return sortedFiles[self.relBook]      

    def generateShape(self):
        print "generic generate shape invoked" 

    def setColumn(self,colText):
        self.columnNum = colText

    def setText(self, textRef):
        self.contentText = textRef

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
                                                                                             
