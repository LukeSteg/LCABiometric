#!/usr/bin/python

import xlrd
from genericChartFactory import genericChartFactory
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION
from pptx.enum.chart import XL_LABEL_POSITION

class pieChartFactory(genericChartFactory):

    def generateShape(self):
        chart = self.slideRef.shapes.add_chart(XL_CHART_TYPE.PIE,self.x,self.y,self.cx,self.cy,self.chart_data).chart
        self.shapeRef.text = ''
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False

        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END



