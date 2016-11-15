#!/usr/bin/python
 
import xlrd
from genericFactory import genericFactory
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_LABEL_POSITION

class genericChartFactory(genericFactory):

    def setColumn(self, colText):
        self.columnNum = colText

    def setData(self, dataRef):
        self.dataRef = dataRef

    def setShape(self, shapeRef):
        self.shapeRef =  shapeRef

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


