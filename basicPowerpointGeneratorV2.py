import xlrd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


#Reading a book
book = xlrd.open_workbook("ExampleDataset.xlsx")

#Using a book
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))

#Using a sheet
print("----------- Sheet 1 -----------")
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell A4 is {0}".format(sh.cell_value(rowx=4, colx=0)))
for rx in range(sh.nrows):
    print(sh.row(rx))
print("--------------------------------")


#Iterating all rows of all sheets
print("------- All Sheets -------")
for ns in range(book.nsheets):
	print "----------- Sheet",ns,"-----------"
	sh = book.sheet_by_index(ns)
	print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
	for rx in range(sh.nrows):
	    print(sh.row(rx))
	print("--------------------------------")

aggregatedSheet = book.sheet_by_index(25)


# create presentation with 1 slide ------
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

# define chart data ---------------------
chart_data = ChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Series 1', (19.2, 21.4, 16.7))

peopleAtGoal = 0
peopleNotAtGoal = 0
numberOfPeople = aggregatedSheet.nrows - 1 
for i in range(aggregatedSheet.nrows - 1):
    numberOfPeople += 1
    if( aggregatedSheet.cell_value(rowx = i + 1, colx = 5) == "Goal"):
        peopleAtGoal += 1;
        print("here" , peopleAtGoal , i );
    else:
        peopleNotAtGoal += 1

percentageAtGoal = 100*(float(peopleAtGoal)/numberOfPeople)
percentageNotAtGoal = 100*(float(peopleNotAtGoal)/numberOfPeople)
print(percentageAtGoal,percentageNotAtGoal)
chart_data = ChartData()
chart_data.categories = ['Goal','Less than Goal']
chart_data.add_series('Series 1',(percentageAtGoal, percentageNotAtGoal))

# add chart to slide --------------------
x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
    )

prs.save('chart-01.pptx')

def getParticipationSlide(presentation, numberOfParticipants, percentageFemale, averageAge)
    participationSlide = presentation.slide_layouts[1]
    titleText = "Who Participated?"
    titleTextFrame.text = titleText

    contentTextFrame.text = """%i Individuals completed fasting biometric screening\nFemale Participants: %i\%\nMale
    Participants: %i\%\nAverage Age: %i years""" %(numberOfParticipants, percentageFemale, 100 - percentageFemale,
    averageAge)

    
