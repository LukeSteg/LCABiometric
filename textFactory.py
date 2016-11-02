#!/usr/bin/python

import xlrd

class textFactory:
    #self.x = None
    #self.y = None
    #self.cx = None
    #self.cy = None 
    #self.textRef = None 
    #self.outputText = None
    #self.outputShape = None
    #self.shapeRef = None
    #self.slideRef = None

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = shapeRef.top
        self.y = shapeRef.left
        self.cx = shapeRef.width
        self.cy = shapeRef.height

    
    def generateText(self):
        return self.outputText

    def generateShape(self):
        newShape = self.slideRef.shapes.add_textbox(self.x, self.y, self.cx, self.cy)
        newShape.text = self.contentText

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

