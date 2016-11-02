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

def parseSlide(slide,slideRef):
    for shape in slide.shapes:
        if shape.has_text_frame:
            parse(slide,shape,shape,slideRef)
            
def createOutput(outputFileName, inputTemplateFileName, DataSheetList):

	prs = Presentation(inputTemplateFileName)

        for slide in prs.slides:
            parseSlide(slide, DataSheetList[0])

	prs.save(outputFileName)


