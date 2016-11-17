#!/usr/bin/python

import xlrd
from genericFactory import genericFactory
from genericFactory import DATA_SHEET_NUM

class textFactory(genericFactory):

    def __init__(self, slideRef, shapeRef):
        super(self.__class__,self).__init__(slideRef, shapeRef)
        self.outputVarType = 'percent'
        self.outputText = 'NO OUTPUT TEXT CREATED'

    def generateShape(self):
        queryStartIndex = self.shapeRef.text.index('#{')
        queryEndIndex = self.shapeRef.text.index('}')
        self.shapeRef.text = self.shapeRef.text[0:queryStartIndex] + self.outputText + self.shapeRef.text[queryEndIndex+1: len(self.shapeRef.text)]
        print self.shapeRef.text

    def setColumn(self,colText):
        self.columnNum = colText 
    
    def setText(self, textRef):
        self.contentText = textRef

    def setOutputVarType(self, outputVarType):
        self.outputVarType = outputVarType

    def computeOutputVar(self, outputVar):
        print 'output var ' + str(outputVar)
        print 'categories ' + str(self.categories)
        if outputVar in self.categories:
            outputVarIndex = self.categories.index(outputVar)
            outputVarCount = self.categoryCount[outputVarIndex]

            if self.outputVarType == 'percent':
                self.outputText = str(int(float(outputVarCount)/float(self.numberOfDataPoints)*100)) + '%'
            elif self.outputVarType == 'count':
                self.outputText = str(outputVarCount)
            else:
                print 'ERROR self.outputVarType not expected varType'
        else:
            print 'ERROR outputVar for text insert was not in detected in given data'


    def getDataFromColumn(self, colNum, fileRef):
        book = xlrd.open_workbook(fileRef);
        dataSheet = book.sheet_by_index(DATA_SHEET_NUM)
        self.chart_data = ChartData()
        rawData = []
        self.categories = []
        self.categoryCount = []
        self.numberOfDataPoints = dataSheet.nrows;

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

