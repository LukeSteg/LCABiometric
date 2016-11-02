#!/usr/bin/python

import xlrd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

class pieChartFactory:

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = shapeRef.top
        self.y = shapeRef.left
        self.cx = shapeRef.width
        self.cy = shapeRef.height


    def generateShape(self):
        print self.x
        print self.y
        print self.cx
        print self.cy
        print self.chart_data
        newShape = self.slideRef.shapes.add_chart(XL_CHART_TYPE.PIE,self.x,self.y,self.cx,self.cy,self.chart_data).chart

    def setColumn(self, colText):
        self.columnNum = colText 
    
    def setData(self, dataRef):
        self.dataRef = dataRef

    def setShape(self, shapeRef):
        self.shapeRef =  shapeRef

    def setX(self, x):
        self.x = x 

    def setY(self, y):
        self.y = y 

    def setCX(self, cx):
        self.cx = cx 

    def setCY(self, cy):
        self.cy = cy 

    def getDataFromColumn(self, colNum, fileRef):
        book = xlrd.open_workbook(fileRef);
        dataSheet = book.sheet_by_index(25)
        self.chart_data = ChartData()
        rawData = []
        categories = []
        categoryCount = []

        for i in range(dataSheet.nrows - 1):
            rawData.append(dataSheet.cell_value(rowx = i + 1, colx = colNum))

        categories = list(set(rawData))

        for i in range(len(categories)):
            categoryCount.append(sum(dataEntry == categories[i] for dataEntry in rawData))
            
        print "categories : "
        print categories
        print "tuple(categorCount) : "
        print tuple(categoryCount)
        self.chart_data.categories = categories
        self.chart_data.add_series('Series 1',tuple(categoryCount)) 


