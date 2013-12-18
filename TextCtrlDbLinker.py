#!/usr/bin/env python
# -*- coding: utf8 -*-

import wx


class TextCtrlDbLinker(wx.TextCtrl):
	def __init__(self):
		pre = wx.PreTextCtrl()
		self.PostCreate(pre)
		self.Bind(wx.EVT_WINDOW_CREATE, self.on_create)
		self.Bind(wx.EVT_KILL_FOCUS, self.on_focus_lost) 


	def on_create(self, event):
		self.Unbind(wx.EVT_WINDOW_CREATE)
		#further initialization goes here


	def on_focus_lost(self, event):
		print self.GetValue()
		
		
