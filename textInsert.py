#!/usr/bin/python

import xlrd

class textFacory:
        self.x = 0
        self.y = 0
        self.cx = 0
        self.cy = 0
        self.textRef = "not set"
        self.outputText = "NULL OUTPUT"
        self.outputShape = None

    def __init__ (self):
        self.textRef = "not set"

    def __init__ (self, excelFilesRef):
        self.excelFiles = excelFilesRef

   def reset():
        self.x = 0
        self.y = 0
        self.cx = 0
        self.cy = 0
        self.textRef = "not set"
        self.outputText = "NULL OUTPUT"

    def generateText():
        return self.outputText

    def generateShape():
        return self.outputShape

    def setColumn(colText):
        self.columnNum = colText 
    
    def setText(textRef):
        self.contentText = textRef

    def setX(x):
        self.x = x 

    def setY(y):
        self.y = y 

    def setCX(cx):
        self.cx = cx 

    def setCY(CY):
        self.CY = CY 

       
