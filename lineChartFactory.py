#!/usr/bin/python

import xlrd
import collections
from genericChartFactory import genericChartFactory
from genericFactory import AGGREGATE_SHEET_NAME
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

class lineChartFactory(genericChartFactory):

    def __init__(self, slideRef, shapeRef):
        super(lineChartFactory, self).__init__(slideRef, shapeRef)
        self.numberOfBooks = 2
        self.chart_data = ChartData()
        self.columnName = 'NO NAME SET'

    def generateShape(self):
        print 'line chart data', self.chart_data
        chart = self.slideRef.shapes.add_chart(XL_CHART_TYPE.LINE,self.x,self.y,self.cx,self.cy,self.chart_data).chart
        self.shapeRef.text = ''
        chart.has_legend = True
        chart.legend.include_in_layout = False
    
    def setNumberOfBooks(self, books):
        self.numberOfBooks = books

    def setColumnName(self, colName):
        self.columnName = colName

    def getColumnNumberByName(self, columnName, sheet):
        columnNumber = 0
        readColumnText = sheet.cell_value(rowx = 0, colx = columnNumber)
        while((readColumnText.strip().upper() != columnName) and (readColumnText != '')):
            print readColumnText
            columnNumber += 1
            readColumnText = sheet.cell_value(rowx = 0, colx = columnNumber)

        if readColumnText == '':
            print 'WARNING, column title: ', readColumnText,  ' not found'

        print 'colnum ',columnNumber ,' readColumnText ', readColumnText

        return columnNumber


    def getDataFromColumn(self, fileDict):
        #book = xlrd.open_workbook(fileRef);
        #dataSheet = book.sheet_by_index(AGGREGATE_SHEET_NAME)#throw error if nothing is returned?
        #self.chart_data = ChartData()
        fileReferences = []
        originalRelativeBook = self.relBook
        
        #TODO make sure enough books have been passed in
        print 'bknum',self.numberOfBooks
        if (self.numberOfBooks > len(fileDict.keys())):
            self.numberOfBooks = len(fileDict.keys())
        for i in range(self.numberOfBooks):
            self.relBook = i + originalRelativeBook
            fileReferences.append(self.getFileFromDict(fileDict))
        
        self.relBook = originalRelativeBook

        allData = []
        dataBySheet = {}
        sheetDict = {}
        categories = []
        #get series titles
        #sum series on a per year/seriesNames basis
        for filePath in fileReferences:
            print 'getting file ', filePath
            book = xlrd.open_workbook(filePath);
            sheetDict[filePath] = book.sheet_by_name(AGGREGATE_SHEET_NAME)
            dataBySheet[filePath] = []
            
            categories.append(fileDict[filePath][1])#gets the year of the file
            colNum = self.getColumnNumberByName(self.columnName, sheetDict[filePath]) 
                
            print 'getting data ', sheetDict[filePath].nrows - 1
            for i in range(sheetDict[filePath].nrows - 1):
                allData.append(sheetDict[filePath].cell_value(rowx = i + 1, colx = colNum))
                print sheetDict[filePath].cell_value(rowx = i + 1, colx = colNum)
                dataBySheet[filePath].append(sheetDict[filePath].cell_value(rowx = i + 1, colx = colNum))

        seriesNames = []
        seriesNames = list(set(allData))
        self.chart_data.categories = categories
        print "categories : "
        print categories
 
        seriesData = collections.defaultdict(list)
        for i in range(len(seriesNames)):
            for filePath in fileReferences:
                seriesData[seriesNames[i]].append(sum(dataEntry == seriesNames[i] for dataEntry in dataBySheet[filePath]))
           
            self.chart_data.add_series(seriesNames[i],tuple(seriesData[seriesNames[i]]))

            print "series: "
            print seriesNames[i],tuple(seriesData[seriesNames[i]])
           


