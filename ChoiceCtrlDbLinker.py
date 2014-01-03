#!/usr/bin/env python
# -*- coding: utf8 -*-

import wx
import Database as db


class ChoiceCtrlDbLinker(wx.Choice):
	def __init__(self):
		pre = wx.PreChoice()
		self.PostCreate(pre)
		self.Bind(wx.EVT_WINDOW_CREATE, self.on_create)
		self.Bind(wx.EVT_KILL_FOCUS, self.on_focus_lost)


	def on_create(self, event):
		self.Unbind(wx.EVT_WINDOW_CREATE)
		#further initialization goes here


	def on_focus_lost(self, event):
		#parse our the table name and field from the control name
		table = '.'.join(self.GetName().split(':')[1].split('.')[0:2])
		table_id = wx.GetTopLevelParent(self).id
		
		field = self.GetName().split(':')[1].split('.')[2]
		new_value = self.GetStringSelection()
		
		if new_value == '':
			new_value = None
		
		db.update_order(table, table_id, field, new_value)

		parent_frame = wx.GetTopLevelParent(self)
		parent_frame.Freeze()
		parent_frame.reset_all()
		parent_frame.populate_all()
		parent_frame.Thaw()
