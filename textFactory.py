#!/usr/bin/python

import xlrd
from genericFactory import genericFactory
from genericFactory import AGGREGATE_SHEET_NAME
from pptx.chart.data import ChartData
from pptx.util import Pt

class textFactory(genericFactory):

    def __init__(self, slideRef, shapeRef):
        super(textFactory,self).__init__(slideRef, shapeRef)
        self.outputVarType = 'PERCENT'
        self.outputText = 'NO OUTPUT TEXT CREATED'
        self.textSize = -1

    def generateShape(self):
        queryStartIndex = self.shapeRef.text.index('#{')
        queryEndIndex = self.shapeRef.text.index('}')
        self.shapeRef.text = self.shapeRef.text[0:queryStartIndex] + self.outputText + self.shapeRef.text[queryEndIndex+1: len(self.shapeRef.text)]

        if(self.textSize != -1):
            for paragraph in self.shapeRef.paragraphs:
                for run in paragraph.runs:
                    run.font = Pt(self.textSize)

        print self.shapeRef.text

    def setText(self, textRef):
        self.contentText = textRef.upper()

    def setTextSize(self, size):
        self.textSize = size

    def setOutputVarType(self, outputVarType):
        if(outputVarType == 'PERCENT' or outputVarType == 'COUNT'):
            self.outputVarType = outputVarType
        else:
            print 'WARNING unexpected \'PERCENTORCOUNT\' value, expected percent or count, found ', outputVar  

    def computeOutputVar(self, outputVar):
        print 'output var ' + str(outputVar)
        print 'categories ' + str(self.categories)
        participantCountString = '#COUNT'
        averageValueOfCategory = '#AVERAGE'

        #get the number of data points 
        if(outputVar.upper() == participantCountString):
            self.outputText = str(self.numberOfDataPoints)

        #get the average of the floats in the data
        if(outputVar.upper() == averageValueOfCategory):
            values = []
            for i in range(len(self.categories)):
                for j in range(len(self.categoryCount)):
                    try:
                        floatValue = float(self.categories[i])
                        values.append(self.categories[i])
                    except (TypeError, ValueError):
                        print 'WARNING, expected to find an int while performing average operation, found: ', self.categories[i]
            self.outputText = str(sum(values)/len(values))

        #get the percentage or count of a var in the data
        elif outputVar in self.categories:
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

