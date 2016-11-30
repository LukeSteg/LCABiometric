#!/usr/bin/python
 
import xlrd
from genericFactory import genericFactory
from genericFactory import AGGREGATE_SHEET_NAME
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_LABEL_POSITION

AGGREGATE_SHEET_NAME = 'Aggregate'

class genericChartFactory(genericFactory):

    def __init__(self, slideRef, shapeRef):
        super(genericChartFactory, self).__init__(slideRef, shapeRef)
        self.titleText = ''

    def setTitle(self, titleText):
        self.titleText = titleText

    def setData(self, dataRef):
        self.dataRef = dataRef

    def setShape(self, shapeRef):
        self.shapeRef =  shapeRef

    def getDataFromColumn(self, fileRef):
        book = xlrd.open_workbook(fileRef);
        dataSheet = book.sheet_by_name(AGGREGATE_SHEET_NAME)
        self.chart_data = ChartData()
        rawData = []
        categories = []
        categoryCount = []
        colNum = self.colNum

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


