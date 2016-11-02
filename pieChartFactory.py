#!/usr/bin/python

import xlrd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

class pieChartFactory:

    def __init__ (self, slideRef, shapeRef):
        self.shapeRef = shapeRef
        self.slideRef = slideRef
        self.x = inputShape.top
        self.y = inputShape.left
        self.cx = inputShape.width
        self.cy = inputShape.height


    def __init__ (self, excelFilesRef):
        self.excelFiles = excelFilesRef
    
    def generatePieChart():
        self.slideRef.shapes.add_chart(XL_CHART_TYPE.PIE,self.x,self.y,self.cx,self.cy,self.chart_data)
        return self.outputText

    def generateShape():
        newShape = self.slideRef.shapes.add_textbox(self.x, self.y, self.cx, self.cy)
        newShape.text = textRef

    def setColumn(colText):
        self.columnNum = colText 
    
    def setData(dataRef):
        self.dataRef = dataRef

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

    def getDataFromColumn(colNum, fileRef):
        book = xlrd.open_workbook(fileRef);
        dataSheet = book.sheet_by_index(25)
        self.chart_data = ChartData()

        rawData = []
        categories = []
        categoryCount = []

        for i in range(aggregatedSheet.nrows - 1):
            rawData.append(aggregatedSheet.cell_value(rowx = i + 1, colx = colNum) 

        catagories = list(set(rawData))

        for(i in range(0:len(catagories)):
            categoryCount.append(dataEntry == categories[i] for dataEntry in rawData))
            
        self.chart_data.categories = categories
        self.chart_data.addSeries('Series 1',tuple(categoryCount)) 


