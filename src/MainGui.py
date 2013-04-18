# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Oct  8 2012)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc

###########################################################################
## Class MainFrame
###########################################################################

class MainFrame ( wx.Frame ):
	
	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"ToC Parser", pos = wx.DefaultPosition, size = wx.Size( -1,-1 ), style = wx.CAPTION|wx.CLOSE_BOX|wx.MAXIMIZE_BOX|wx.MINIMIZE_BOX|wx.SYSTEM_MENU|wx.TAB_TRAVERSAL )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		
		bSizer1 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_panel1 = wx.Panel( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL )
		bSizer15 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer6 = wx.StaticBoxSizer( wx.StaticBox( self.m_panel1, wx.ID_ANY, u"Destination file" ), wx.VERTICAL )
		
		self.m_filePickerDest = wx.FilePickerCtrl( self.m_panel1, wx.ID_ANY, wx.EmptyString, u"Select a file", u"Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", wx.DefaultPosition, wx.DefaultSize, wx.FLP_OVERWRITE_PROMPT|wx.FLP_SAVE|wx.FLP_USE_TEXTCTRL )
		sbSizer6.Add( self.m_filePickerDest, 1, wx.ALIGN_CENTER_HORIZONTAL|wx.EXPAND, 5 )
		
		
		bSizer15.Add( sbSizer6, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.EXPAND|wx.BOTTOM, 5 )
		
		bSizer5 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.m_checkBoxRememberFilename = wx.CheckBox( self.m_panel1, wx.ID_ANY, u"Remember filename", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer5.Add( self.m_checkBoxRememberFilename, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		self.m_staticline5 = wx.StaticLine( self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL )
		bSizer5.Add( self.m_staticline5, 0, wx.EXPAND|wx.TOP|wx.BOTTOM|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		self.m_checkBoxConvertUnnumbered = wx.CheckBox( self.m_panel1, wx.ID_ANY, u"Convert unnumbered sections", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer5.Add( self.m_checkBoxConvertUnnumbered, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		
		bSizer15.Add( bSizer5, 0, wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.m_staticline21 = wx.StaticLine( self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL )
		bSizer15.Add( self.m_staticline21, 0, wx.EXPAND, 5 )
		
		bSizer20 = wx.BoxSizer( wx.HORIZONTAL )
		
		self.m_buttonConvertExcel = wx.Button( self.m_panel1, wx.ID_ANY, u"Convert To Excel", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer20.Add( self.m_buttonConvertExcel, 0, wx.ALL, 5 )
		
		self.m_staticline3 = wx.StaticLine( self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL )
		bSizer20.Add( self.m_staticline3, 0, wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		self.m_buttonAbout = wx.Button( self.m_panel1, wx.ID_ANY, u"About", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer20.Add( self.m_buttonAbout, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		self.m_staticline31 = wx.StaticLine( self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_VERTICAL )
		bSizer20.Add( self.m_staticline31, 0, wx.EXPAND|wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		self.m_buttonClearRegistry = wx.Button( self.m_panel1, wx.ID_ANY, u"Clear Registry && Exit", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer20.Add( self.m_buttonClearRegistry, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 5 )
		
		
		bSizer15.Add( bSizer20, 0, wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.m_staticline2 = wx.StaticLine( self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL )
		bSizer15.Add( self.m_staticline2, 0, wx.EXPAND, 5 )
		
		self.m_staticText9 = wx.StaticText( self.m_panel1, wx.ID_ANY, u"Copy the ToC from the Word document and paste it in the field below", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText9.Wrap( -1 )
		self.m_staticText9.SetFont( wx.Font( wx.NORMAL_FONT.GetPointSize(), 70, 90, 92, False, wx.EmptyString ) )
		
		bSizer15.Add( self.m_staticText9, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.m_textCtrlToC = wx.TextCtrl( self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 500,400 ), wx.TE_DONTWRAP|wx.TE_MULTILINE )
		bSizer15.Add( self.m_textCtrlToC, 1, wx.EXPAND|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		sbSizer3 = wx.StaticBoxSizer( wx.StaticBox( self.m_panel1, wx.ID_ANY, u"Conversion Log" ), wx.VERTICAL )
		
		self.m_textCtrlConvLog = wx.TextCtrl( self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( -1,100 ), wx.TE_DONTWRAP|wx.TE_MULTILINE )
		sbSizer3.Add( self.m_textCtrlConvLog, 0, wx.EXPAND, 5 )
		
		
		bSizer15.Add( sbSizer3, 0, wx.EXPAND, 5 )
		
		
		self.m_panel1.SetSizer( bSizer15 )
		self.m_panel1.Layout()
		bSizer15.Fit( self.m_panel1 )
		bSizer1.Add( self.m_panel1, 1, wx.ALIGN_CENTER_HORIZONTAL|wx.EXPAND, 5 )
		
		
		self.SetSizer( bSizer1 )
		self.Layout()
		bSizer1.Fit( self )
		self.m_statusBar = self.CreateStatusBar( 1, wx.ST_SIZEGRIP, wx.ID_ANY )
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.m_filePickerDest.Bind( wx.EVT_FILEPICKER_CHANGED, self.OnFileChangedDestFile )
		self.m_checkBoxRememberFilename.Bind( wx.EVT_CHECKBOX, self.OnCheckBoxRememberFilename )
		self.m_checkBoxConvertUnnumbered.Bind( wx.EVT_CHECKBOX, self.OnCheckBoxConvertUnnumbered )
		self.m_buttonConvertExcel.Bind( wx.EVT_BUTTON, self.OnButtonClickConvertToExcel )
		self.m_buttonAbout.Bind( wx.EVT_BUTTON, self.OnButtonClickAbout )
		self.m_buttonClearRegistry.Bind( wx.EVT_BUTTON, self.OnButtonClickClearRegistryAndExit )
		self.m_textCtrlToC.Bind( wx.EVT_TEXT, self.OnTextToC )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def OnFileChangedDestFile( self, event ):
		event.Skip()
	
	def OnCheckBoxRememberFilename( self, event ):
		event.Skip()
	
	def OnCheckBoxConvertUnnumbered( self, event ):
		event.Skip()
	
	def OnButtonClickConvertToExcel( self, event ):
		event.Skip()
	
	def OnButtonClickAbout( self, event ):
		event.Skip()
	
	def OnButtonClickClearRegistryAndExit( self, event ):
		event.Skip()
	
	def OnTextToC( self, event ):
		event.Skip()
	

###########################################################################
## Class AboutDialog
###########################################################################

class AboutDialog ( wx.Dialog ):
	
	def __init__( self, parent ):
		wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = u"About ToC Parser", pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_DIALOG_STYLE )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		
		bSizer14 = wx.BoxSizer( wx.VERTICAL )
		
		sbSizer7 = wx.StaticBoxSizer( wx.StaticBox( self, wx.ID_ANY, wx.EmptyString ), wx.VERTICAL )
		
		self.m_staticText5 = wx.StaticText( self, wx.ID_ANY, u"Convert the a Word Table of Contents\nto an Excel (R) File", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTRE )
		self.m_staticText5.Wrap( -1 )
		sbSizer7.Add( self.m_staticText5, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		self.m_staticline10 = wx.StaticLine( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL )
		sbSizer7.Add( self.m_staticline10, 0, wx.EXPAND |wx.ALL, 5 )
		
		self.m_staticText51 = wx.StaticText( self, wx.ID_ANY, u"(C) 2012 Daniel Rodriguez", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText51.Wrap( -1 )
		sbSizer7.Add( self.m_staticText51, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		bSizer14.Add( sbSizer7, 1, wx.EXPAND|wx.ALL, 5 )
		
		self.m_button7 = wx.Button( self, wx.ID_OK, u"Ok", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer14.Add( self.m_button7, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		
		
		self.SetSizer( bSizer14 )
		self.Layout()
		bSizer14.Fit( self )
		
		self.Centre( wx.BOTH )
		
		# Connect Events
		self.m_button7.Bind( wx.EVT_BUTTON, self.OnButtonClickOk )
	
	def __del__( self ):
		pass
	
	
	# Virtual event handlers, overide them in your derived class
	def OnButtonClickOk( self, event ):
		event.Skip()
	

