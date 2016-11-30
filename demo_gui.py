#!/usr/bin/env python
import datetime
import wx
import os
import fnmatch
import templateReaderV3 as trV3

class FileSelectionFrame(wx.Frame):
	def __init__(self, parent):
		super(FileSelectionFrame, self).__init__(None, title = "PowerPoint Generator v0.1", size = (1000, 400))

class FileSelectionPanel(wx.Panel):
	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		
		self.template_filename = ""
		self.template_dir = "."
		
		self.output_filename = ""
		self.output_dir = "."
		
		self.excel_dir = "."
		self.excel_dir_set = False
		self.excel_list_all = {}
		self.excel_list_dates = {}
		self.excel_dict_use = {}
		
		self.TFB_pressed = False
		self.OFB_pressed = False
		self.EDB_pressed = False
		
		self.last_dir = "."
		
		# get template file
		self.template_file_button = wx.Button(self, label = "Choose Type of Template for Report")
		self.Bind(wx.EVT_BUTTON, self.onTFB, self.template_file_button)
		self.TFB_label = wx.TextCtrl(self, size=(500,-1), style = wx.TE_READONLY)
		
		# get output file
		self.output_file_button = wx.Button(self, label = "Please name your Powerpoint")
		self.Bind(wx.EVT_BUTTON, self.onOFB, self.output_file_button)
		self.OFB_label = wx.TextCtrl(self, size=(500,-1), style = wx.TE_READONLY)
		
		# get directory of input excel file(s)
		self.excel_dir_button = wx.Button(self, label = "Name of Company Folder")
		self.Bind(wx.EVT_BUTTON, self.onEDB, self.excel_dir_button)
		self.EDB_label = wx.TextCtrl(self, size=(500,-1), style = wx.TE_READONLY)
		
		#TO-DO: add checkboxes foreach excel file in the given directory
		#so that the user may opt to use some of them, but not all.
		#ecxel file checkbox list (to be created upon choosing a directory)
		self.excel_dir_checkboxes = {}
		self.excel_dir_datewheels = {}
		
		#run button
		self.run_button = wx.Button(self, label = "Run")
		self.Bind(wx.EVT_BUTTON, self.onRun, self.run_button)
		self.run_button.Enable(False)
		self.setupSizers()
		
		
	def setupSizers(self):
		# create some sizers
		#mainSizer = wx.BoxSizer(wx.HORIZONTAL)
		vertSizer = wx.BoxSizer(wx.VERTICAL)
		file_options = wx.GridBagSizer(hgap = 5, vgap = 5)
		excel_dir_grid = wx.GridBagSizer(hgap = 5, vgap = 5)
		
		#add things to the sizers
		file_options.Add(self.template_file_button, pos = (1,0))
		file_options.Add(self.TFB_label, pos = (1,1))
		
		file_options.Add(self.output_file_button, pos = (2,0))
		file_options.Add(self.OFB_label, pos = (2,1))
		
		file_options.Add(self.excel_dir_button, pos = (3,0))
		file_options.Add(self.EDB_label, pos = (3,1))
		
		i = 1
		for fname in self.excel_list_all:
			box = self.excel_dir_checkboxes[fname]
			tup = self.excel_dir_datewheels[fname]
			month = tup[0]
			year = tup[1]
			excel_dir_grid.Add(box, pos = (i,1))
			excel_dir_grid.Add(month, pos = (i,2))
			excel_dir_grid.Add(year, pos = (i,3))
			i += 1
		
		#finally set up the sizer hierarchy
		vertSizer.Add(file_options, 0, wx.ALL, 5)
		vertSizer.Add(excel_dir_grid, 0, wx.ALL, 5)
		vertSizer.Add(self.run_button, 0, wx.CENTER)
		#mainSizer.Add(vertSizer, 0, wx.ALL, 5)
		#mainSizer.Add(excel_dir_grid, 0, wx.ALL, 5)
		self.SetSizerAndFit(vertSizer)
		
	def tryEnableRunButton(self):
		if self.TFB_pressed and	self.OFB_pressed and self.EDB_pressed:
			self.run_button.Enable(True)	
	
	def defaultFileDialogOptions(self, title = 'Choose a file'):
		''' Return a dictionary with file dialog options that can be
			used in both the save file dialog as well as in the open
			file dialog. '''
		return dict(message=title, defaultDir=self.last_dir, wildcard='*.*')
		
	def onTFB(self, event):
		self.TFB_pressed = True
		answer = self.askUserForFilename(style = wx.FD_OPEN, **self.defaultFileDialogOptions('Choose a Template Powerpoint file'))
		if answer[0]:
			self.template_filename = answer[2]
			self.template_dir = answer[1]
			self.TFB_label.SetValue(str(os.path.join(self.template_dir,self.template_filename)))
			self.tryEnableRunButton()
		
		
	def onOFB(self, event):
		self.OFB_pressed = True
		answer = self.askUserForFilename(style = wx.FD_SAVE, **self.defaultFileDialogOptions('Choose a file to write output to'))
		if answer[0]:
			self.output_filename = answer[2]
			self.output_dir = answer[1]
			self.OFB_label.SetValue(str(os.path.join(self.output_dir,self.output_filename)))
			self.tryEnableRunButton()
		
	def onEDB(self, event):
		self.EDB_pressed = True
		answer = self.askUserForFilename(style = wx.FD_CHANGE_DIR, **self.defaultFileDialogOptions('Choose a folder containing excel files'))
		if answer[0]:
			self.excel_dir = answer[1]
			self.excel_dir_set = True
			self.EDB_label.SetValue(str(self.excel_dir))
			self.excel_list_all = {}

			thisYear = datetime.date.today().year
			thisMonth = datetime.date.today().month
			monthList = ['Jan', 'Feb', 'Mar', 'April', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

			for filename in os.listdir(self.excel_dir):
				if fnmatch.fnmatch(filename, '*.xlsx'):
					self.excel_list_all[filename] = False
					self.excel_list_dates[filename] = (monthList[thisMonth-1], thisYear)
			self.Unbind(wx.EVT_CHECKBOX)
			self.excel_dir_grid = wx.BoxSizer(wx.VERTICAL)
			self.excel_dir_checkboxes = {}
			for fname in self.excel_list_all:
				checkbox = wx.CheckBox(self, label=fname)
				self.Bind(wx.EVT_CHECKBOX, self.onCheck, checkbox)
				self.excel_dir_checkboxes[fname] = checkbox
				


				numYears = 15
				yearList = [str(x) for x in range(thisYear, thisYear - numYears, -1)]
				
				month = wx.ComboBox(self, size=(90, -1), choices=monthList, style=wx.CB_DROPDOWN)
				year = wx.ComboBox(self, size=(90, -1), choices=yearList, style=wx.CB_DROPDOWN)
				month.SetSelection(thisMonth-1)
				year.SetSelection(0)
				self.Bind(wx.EVT_COMBOBOX, self.onCheck, month)
				self.Bind(wx.EVT_COMBOBOX, self.onCheck, year)
				self.excel_dir_datewheels[fname] = (month, year)
				
			self.setupSizers()
			self.tryEnableRunButton()
				
	def onCheck(self, event):
		for fname, box in self.excel_dir_checkboxes.items():
			self.excel_list_all[fname] = box.GetValue()
			tup = self.excel_dir_datewheels[fname]
			month = tup[0].GetValue()
			year = tup[1].GetValue()
			self.excel_list_dates[fname] = (month, year)
		
	def onRun(self, event):
		self.excel_dict_use = {self.excel_dir + os.sep + fname : (self.excel_list_dates[fname][0],self.excel_list_dates[fname][1]) for fname, val in self.excel_list_all.items() if val}
		trV3.createOutput(self.output_dir + os.sep + self.output_filename, self.template_dir + os.sep + self.template_filename, self.excel_dict_use)
		raise SystemExit
		
	def askUserForFilename(self, **dialogOptions):
		dialog = wx.FileDialog(self, **dialogOptions)
		filename = "wasd.txt"
		dirname = "."
		if dialog.ShowModal() == wx.ID_OK:
			userProvidedFilename = True
			filename = dialog.GetFilename()
			dirname = dialog.GetDirectory()
			self.last_dir = dirname
		else:
			userProvidedFilename = False
		dialog.Destroy()
		return [userProvidedFilename, dirname, filename]
	
app = wx.App(False)
frame = FileSelectionFrame(None)
panel = FileSelectionPanel(frame)
frame.Show()
app.MainLoop()	
	
