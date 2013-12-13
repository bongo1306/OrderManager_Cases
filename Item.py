


#!/usr/bin/env python
# -*- coding: utf8 -*-

import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import datetime as dt
import General as gn
import Database as db


class ItemFrame(wx.Frame):
	def __init__(self, parent, id):
		#load frame XRC description
		pre = wx.PreFrame()
		res = xrc.XmlResource.Get() 
		res.LoadOnFrame(pre, parent, "frame:item") 
		self.PostCreate(pre)
		self.SetIcon(wx.Icon(gn.resource_path('OrderManager.ico'), wx.BITMAP_TYPE_ICO))
		
		self.id = id
		
		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)

		#misc
		self.SetTitle('Item ID {}'.format(self.id))
		self.SetSize((800, 600))
		self.Center()


		self.Show()


	def init_changes_tab(self):
		pass
		
	def reset_changes_tab(self):
		pass

	def populate_changes_tab(self):
		pass


	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()
