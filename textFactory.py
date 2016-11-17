#!/usr/bin/python

import xlrd
from genericFactory import genericFactory
from genericFactory import DATA_SHEET_NUM

class textFactory(genericFactory):

    def generateShape(self):
        queryStartIndex = self.shapeRef.text.index('#{')
        queryEndIndex = self.shapeRef.text.index('}')
        self.shapeRef.text = self.shapeRef.text[0:queryStartIndex] + self.contentText + self.shapeRef.text[queryEndIndex+1: len(self.shapeRef.text)]
        print self.shapeRef.text

    def setColumn(self,colText):
        self.columnNum = colText 
    
    def setText(self, textRef):
        self.contentText = textRef

    def getDataFromColumn(self, colNum, fileRef):
        book = xlrd.open_workbook(fileRef);
        dataSheet = book.sheet_by_index(DATA_SHEET_NUM)
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

