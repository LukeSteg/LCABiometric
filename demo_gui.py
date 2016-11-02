#!/usr/bin/env python

import wx
import os
import fnmatch
import templateReaderV3 as trV3

class FileSelectionFrame(wx.Frame):
	def __init__(self, parent):
		super(FileSelectionFrame, self).__init__(None, title = "PowerPoint Generator v0.1", size = (1000, 230))

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
		self.excel_list_use = []
		
		self.last_dir = "."
		
		# get template file
		self.template_file_button = wx.Button(self, label = "Choose Template Powerpoint File")
		self.Bind(wx.EVT_BUTTON, self.onTFB, self.template_file_button)
		self.TFB_label = wx.TextCtrl(self, size=(500,-1), style = wx.TE_READONLY)
		
		# get output file
		self.output_file_button = wx.Button(self, label = "Choose Output Filename")
		self.Bind(wx.EVT_BUTTON, self.onOFB, self.output_file_button)
		self.OFB_label = wx.TextCtrl(self, size=(500,-1), style = wx.TE_READONLY)
		
		# get directory of input excel file(s)
		self.excel_dir_button = wx.Button(self, label = "Choose Directory of Input Excel File(s)")
		self.Bind(wx.EVT_BUTTON, self.onEDB, self.excel_dir_button)
		self.EDB_label = wx.TextCtrl(self, size=(500,-1), style = wx.TE_READONLY)
		
		#TO-DO: add checkboxes foreach excel file in the given directory
		#so that the user may opt to use some of them, but not all.
		#ecxel file checkbox list (to be created upon choosing a directory)
		self.excel_dir_checkboxes = {}
		
		#run button
		self.run_button = wx.Button(self, label = "Run")
		self.Bind(wx.EVT_BUTTON, self.onRun, self.run_button)
		self.setupSizers()
		
		
	def setupSizers(self):
		# create some sizers
		mainSizer = wx.BoxSizer(wx.HORIZONTAL)
		vertSizer = wx.BoxSizer(wx.VERTICAL)
		file_options = wx.GridBagSizer(hgap = 5, vgap = 5)
		excel_dir_grid = wx.BoxSizer(wx.VERTICAL)
		
		#add things to the sizers
		file_options.Add(self.template_file_button, pos = (1,0))
		file_options.Add(self.TFB_label, pos = (1,1))
		
		file_options.Add(self.output_file_button, pos = (2,0))
		file_options.Add(self.OFB_label, pos = (2,1))
		
		file_options.Add(self.excel_dir_button, pos = (3,0))
		file_options.Add(self.EDB_label, pos = (3,1))
		
		for box in self.excel_dir_checkboxes.values():
			excel_dir_grid.Add(box)
		
		#finally set up the sizer hierarchy
		vertSizer.Add(file_options, 0, wx.ALL, 5)
		vertSizer.Add(self.run_button, 0, wx.CENTER)
		mainSizer.Add(vertSizer, 0, wx.ALL, 5)
		mainSizer.Add(excel_dir_grid, 0, wx.ALL, 5)
		self.SetSizerAndFit(mainSizer)
		
		
	def defaultFileDialogOptions(self, title = 'Choose a file'):
		''' Return a dictionary with file dialog options that can be
			used in both the save file dialog as well as in the open
			file dialog. '''
		return dict(message=title, defaultDir=self.last_dir, wildcard='*.*')
		
	def onTFB(self, event):
		answer = self.askUserForFilename(style = wx.OPEN, **self.defaultFileDialogOptions('Choose a Template Powerpoint file'))
		if answer[0]:
			self.template_filename = answer[2]
			self.template_dir = answer[1]
			self.TFB_label.SetValue(str(os.path.join(self.template_dir,self.template_filename)))
		
		
	def onOFB(self, event):
		answer = self.askUserForFilename(style = wx.SAVE, **self.defaultFileDialogOptions('Choose a file to write output to'))
		if answer[0]:
			self.output_filename = answer[2]
			self.output_dir = answer[1]
			self.OFB_label.SetValue(str(os.path.join(self.output_dir,self.output_filename)))
		
	def onEDB(self, event):
		answer = self.askUserForFilename(style = wx.CHANGE_DIR, **self.defaultFileDialogOptions('Choose a folder containing excel files'))
		if answer[0]:
			self.excel_dir = answer[1]
			self.excel_dir_set = True
			self.EDB_label.SetValue(str(self.excel_dir))
			self.excel_list_all = {}
			for filename in os.listdir(self.excel_dir):
				if fnmatch.fnmatch(filename, '*.xlsx'):
					self.excel_list_all[filename] = False
			self.Unbind(wx.EVT_CHECKBOX)
			self.excel_dir_grid = wx.BoxSizer(wx.VERTICAL)
			self.excel_dir_checkboxes = {}
			for fname in self.excel_list_all:
				checkbox = wx.CheckBox(self, label=fname)
				self.Bind(wx.EVT_CHECKBOX, self.onCheck, checkbox)
				self.excel_dir_checkboxes[fname] = checkbox
			self.setupSizers()
				
	def onCheck(self, event):
		for fname, box in self.excel_dir_checkboxes.items():
			self.excel_list_all[fname] = box.GetValue()
		
	def onRun(self, event):
		self.excel_list_use = [self.excel_dir + os.sep + fname for fname, val in self.excel_list_all.items() if val]
		trV3.createOutput(self.output_dir + os.sep + self.output_filename, self.template_dir + os.sep + self.template_filename, self.excel_list_use)
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
	
