#!/usr/bin/python

import xlrd
from genericFactory import genericFactory
from genericFactory import AGGREGATE_SHEET_NAME
from pptx.chart.data import ChartData

class textFactory(genericFactory):

    def __init__(self, slideRef, shapeRef):
        super(textFactory,self).__init__(slideRef, shapeRef)
        self.outputVarType = 'PERCENT'
        self.outputText = 'NO OUTPUT TEXT CREATED'

    def generateShape(self):
        queryStartIndex = self.shapeRef.text.index('#{')
        queryEndIndex = self.shapeRef.text.index('}')
        self.shapeRef.text = self.shapeRef.text[0:queryStartIndex] + self.outputText + self.shapeRef.text[queryEndIndex+1: len(self.shapeRef.text)]
        print self.shapeRef.text

    def setText(self, textRef):
        self.contentText = textRef.upper()

    def setOutputVarType(self, outputVarType):
        self.outputVarType = outputVarType

    def computeOutputVar(self, outputVar):
        print 'output var ' + str(outputVar)
        print 'categories ' + str(self.categories)
        if outputVar in self.categories:
            outputVarIndex = self.categories.index(outputVar)
            outputVarCount = self.categoryCount[outputVarIndex]

            if self.outputVarType == 'PERCENT':
                self.outputText = str(int(float(outputVarCount)/float(self.numberOfDataPoints)*100)) + '%'
            elif self.outputVarType == 'COUNT':
                self.outputText = str(outputVarCount)
            else:
                print 'ERROR self.outputVarType not expected varType'
        else:
            print 'ERROR outputVar for text insert was not in detected in given data'


    def getDataFromColumn(self, fileRef):
        book = xlrd.open_workbook(fileRef);
        dataSheet = book.sheet_by_name(AGGREGATE_SHEET_NAME)#throw error if nothing is returned?
        self.chart_data = ChartData()
        rawData = []
        self.categories = []
        self.categoryCount = []
        self.numberOfDataPoints = dataSheet.nrows;
        colNum = self.colNum

        for i in range(dataSheet.nrows - 1):
            rawData.append(dataSheet.cell_value(rowx = i + 1, colx = colNum))
        
        unicodeCategories = list(set(rawData))
        self.categories = list(set(rawData))
        
        for i in range(len(self.categories)):
            self.categories[i] = str(self.categories[i]).upper()

        for i in range(len(unicodeCategories)):
            self.categoryCount.append(sum(dataEntry == unicodeCategories[i] for dataEntry in rawData))

        print "categories : "
        print self.categories
        print "tuple(categorCount) : "
        print tuple(self.categoryCount)

