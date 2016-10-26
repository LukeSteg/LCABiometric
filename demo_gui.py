#!/usr/bin/env python

import wx
import os.path

class FileSelectionFrame(wx.Frame):
	def __init__(self, parent):
		super(FileSelectionFrame, self).__init__(None, title = "PowerPoint Generator v0.1", size = (900, 230))

class FileSelectionPanel(wx.Panel):
	def __init__(self, parent):
		wx.Panel.__init__(self, parent)
		
		self.template_filename = ""
		self.template_dir = "."
		
		self.output_filename = ""
		self.output_dir = "."
		
		self.excel_dir = "."
		self.excel_dir_set = False
		
		self.last_dir = "."
		
		# create some sizers
		mainSizer = wx.BoxSizer(wx.VERTICAL)
		grid = wx.GridBagSizer(hgap = 5, vgap = 5)
		excel_dir_grid = wx.GridBagSizer(hgap = 5, vgap = 5)
		
		# get template file
		self.template_file_button = wx.Button(self, label = "Choose Template Powerpoint File")
		self.Bind(wx.EVT_BUTTON, self.onTFB, self.template_file_button)
		self.TFB_label = wx.TextCtrl(self, size=(600,-1), style = wx.TE_READONLY)
		
		# get output file
		self.output_file_button = wx.Button(self, label = "Choose Output Filename")
		self.Bind(wx.EVT_BUTTON, self.onOFB, self.output_file_button)
		self.OFB_label = wx.TextCtrl(self, size=(600,-1), style = wx.TE_READONLY)
		
		# get directory of input excel file(s)
		self.excel_dir_button = wx.Button(self, label = "Choose Directory of Input Excel File(s)")
		self.Bind(wx.EVT_BUTTON, self.onEDB, self.excel_dir_button)
		self.EDB_label = wx.TextCtrl(self, size=(600,-1), style = wx.TE_READONLY)
		
		#TO-DO: add checkboxes foreach excel file in the given directory
		#so that the user may opt to use some of them, but not all.
		
		#run button
		self.run_button = wx.Button(self, label = "Run")
		self.Bind(wx.EVT_BUTTON, self.onRun, self.run_button)
		
		#add things to the sizers
		grid.Add(self.template_file_button, pos = (1,0))
		grid.Add(self.TFB_label, pos = (1,1))
		
		grid.Add(self.output_file_button, pos = (2,0))
		grid.Add(self.OFB_label, pos = (2,1))
		
		grid.Add(self.excel_dir_button, pos = (3,0))
		grid.Add(self.EDB_label, pos = (3,1))
		
		#finally set up the sizer hierarchy
		mainSizer.Add(grid, 0, wx.ALL, 5)
		mainSizer.Add(excel_dir_grid, 0, wx.ALL, 5)
		mainSizer.Add(self.run_button, 0, wx.CENTER)
		self.SetSizerAndFit(mainSizer)
		
	def defaultFileDialogOptions(self):
		''' Return a dictionary with file dialog options that can be
			used in both the save file dialog as well as in the open
			file dialog. '''
		return dict(message='Choose a file', defaultDir=self.last_dir, wildcard='*.*')
		
	def onTFB(self, event):
		answer = self.askUserForFilename(style = wx.OPEN, **self.defaultFileDialogOptions())
		if answer[0]:
			self.template_filename = answer[2]
			self.template_dir = answer[1]
			self.TFB_label.SetValue(str(os.path.join(self.template_dir,self.template_filename)))
		
		
	def onOFB(self, event):
		answer = self.askUserForFilename(style = wx.OPEN, **self.defaultFileDialogOptions())
		if answer[0]:
			self.output_filename = answer[2]
			self.output_dir = answer[1]
			self.OFB_label.SetValue(str(os.path.join(self.output_dir,self.output_filename)))
		
	def onEDB(self, event):
		answer = self.askUserForFilename(style = wx.OPEN, **self.defaultFileDialogOptions())
		if answer[0]:
			self.excel_dir = answer[1]
			self.excel_dir_set = True
			self.EDB_label.SetValue(str(self.excel_dir))
		
	def onRun(self, event):
		pass
		
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
	
