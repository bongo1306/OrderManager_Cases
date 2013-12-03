#!/usr/bin/env python
# -*- coding: utf8 -*-

import wx

FORGET THIS MESS!!! 
instead, create a text control when a static text is clicked for copying text

class LabelStyleTextCtrl(wx.TextCtrl):
	def __init__(self):
		pre = wx.PreTextCtrl()
		self.PostCreate(pre)
		self.Bind(wx.EVT_WINDOW_CREATE, self.on_create)
		self.Bind(wx.EVT_SIZE, self.on_size)
		
		self.previous_size = (9999, 9999)

	def on_create(self, event):
		self.Unbind(wx.EVT_WINDOW_CREATE)
		
		self.SetBackgroundColour(self.GetParent().GetBackgroundColour())
		self.SetBackgroundColour((200, 200, 200))
		self.SetWindowStyle(wx.NO_BORDER)

		self.SetSize((9999, 9999))


	def on_size(self, event):
		print 'on_size:', self.GetSize(), self.previous_size
		
		if self.GetSize()[1] != self.previous_size[1]:
			self.SetSize((self.GetSize()[0], wx.WindowDC(self).GetMultiLineTextExtent(self.GetValue(), self.GetFont())[1]))
			
			self.previous_size = self.GetSize()
			
			print '  SIZING'
			
			self.GetParent().SetBackgroundColour((0, 0, 0))
			
			#self.GetParentFrame().Layout()
			
			print wx.GetTopLevelParent(self).SetBackgroundColour((0, 200, 200))
		

