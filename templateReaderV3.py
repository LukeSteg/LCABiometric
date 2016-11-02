#!/usr/bin/python

import xlrd
from shutil import copyfile
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches,Pt
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_LABEL_POSITION
from Parser import parse

def parseSlide(slide):
    for shape in slide.shapes:
        if shape.has_text_frame:
            parse(slide,shape,shape)
            
def createOutput(outputFileName, inputTemplateFileName, DataSheetList):
#	outputFileName = 'exampleOutput.pptx'
#	inputTemplateFileName = 'exampleTemplate4.pptx'

	prs = Presentation(inputTemplateFileName)

	slide = prs.slides[0]

	slide = prs.slides[1]
	parseSlide(slide)


#	book = xlrd.open_workbook("ExampleDataset.xlsx")
	book = xlrd.open_workbook(DataSheetList[0])
	aggregatedSheet = book.sheet_by_index(25)
	chart_data = ChartData()
	chart_data.categories = ['East', 'West', 'Midwest']
	chart_data.add_series('Series 1', (19.2, 21.4, 16.7))

	peopleAtGoal = 0
	peopleNotAtGoal = 0
	numberOfPeople = 0
	for i in range(aggregatedSheet.nrows - 1):
	    numberOfPeople += 1
	    if( aggregatedSheet.cell_value(rowx = i + 1, colx = 5) == "Goal"):
		peopleAtGoal += 1;
	    else:
		peopleNotAtGoal += 1
		percentageAtGoal = 100*(float(peopleAtGoal)/numberOfPeople)
		percentageNotAtGoal = 100*(float(peopleNotAtGoal)/numberOfPeople)
		chart_data = ChartData()
		chart_data.categories = ['Goal','Less than Goal']
		chart_data.add_series('Series 1',(percentageAtGoal, percentageNotAtGoal))

	slide = prs.slides[2]
	parseSlide(slide)        

	prs.save(outputFileName)


