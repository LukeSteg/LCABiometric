#!/usr/bin/python

import xlrd

DATA_SHEET_NUM = 25

class genericFactory(object):

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = shapeRef.top
        self.y = shapeRef.left
        self.cx = shapeRef.width
        self.cy = shapeRef.height
        self.relYear = 0

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
                                                                                             
