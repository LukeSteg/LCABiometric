import xlrd

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

