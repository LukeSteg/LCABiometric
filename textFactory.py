#!/usr/bin/python

import xlrd

class textFactory:
        self.x = None
        self.y = None
        self.cx = None
        self.cy = None 
        self.textRef = None 
        self.outputText = None
        self.outputShape = None
        self.shapeRef = None
        self.slideRef = None

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = inputShape.top
        self.y = inputShape.left
        self.cx = inputShape.width
        self.cy = inputShape.height


    def __init__ (self, excelFilesRef):
        self.excelFiles = excelFilesRef
    
    def generateText():
        return self.outputText

    def generateShape():
        newShape = self.slideRef.shapes.add_textbox(self.x, self.y, self.cx, self.cy)
        newShape.text = textRef

    def setColumn(colText):
        self.columnNum = colText 
    
    def setText(textRef):
        self.contentText = textRef

    def setShape(shapeRef):
        self.shapeRef =  shapeRef

    def setX(x):
        self.x = x 

    def setY(y):
        self.y = y 

    def setCX(cx):
        self.cx = cx 

    def setCY(CY):
        self.CY = CY

    @staticmethod
    def setCol():

        
        
