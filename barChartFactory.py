#!/usr/bin/python

import xlrd
from genericChartFactory import genericChartFactory
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

class barChartFactory(genericChartFactory):


    def generateShape(self):
        newChart = self.slideRef.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,self.x,self.y,self.cx,self.cy,self.chart_data).chart
        self.shapeRef.text = ''


