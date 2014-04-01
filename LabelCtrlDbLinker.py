#!/usr/bin/env python
# -*- coding: utf8 -*-

import wx
import Database as db


class LabelCtrlDbLinker(wx.StaticText):
	def __init__(self):
		pre = wx.PreStaticText()
		self.PostCreate(pre)
		self.Bind(wx.EVT_WINDOW_CREATE, self.on_create)
		self.Bind(wx.EVT_LEFT_DOWN, self.on_left_down)


	def on_create(self, event):
		self.Unbind(wx.EVT_WINDOW_CREATE)
		#further initialization goes here
		
		self.SetCursor(wx.StockCursor(wx.CURSOR_HAND))


	def on_left_down(self, event):
		LabelEditFrame(self)



class LabelEditFrame(wx.Frame):
	def __init__(self, parent):
		wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size(-1,-1), style = wx.FRAME_FLOAT_ON_PARENT|wx.FRAME_NO_TASKBAR|wx.NO_BORDER)

		self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
		self.SetFont(wx.Font(10, 70, 90, 90, False, wx.EmptyString))
		
		sizer1 = wx.BoxSizer(wx.HORIZONTAL)
		
		self.panel = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
		sizer2 = wx.BoxSizer(wx.HORIZONTAL)
		
		self.text = wx.TextCtrl(self.panel, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(200,-1), wx.TE_PROCESS_ENTER)
		sizer2.Add(self.text, 0, wx.ALL|wx.ALIGN_CENTER_VERTICAL, 0)
		
		self.panel.SetSizer(sizer2)
		self.panel.Layout()
		sizer2.Fit(self.panel)
		sizer1.Add(self.panel, 1, wx.EXPAND, 0)
		
		self.SetSizer(sizer1)
		self.Layout()

		self.SetSize(self.GetBestSize())
		
		self.parent = parent
		
		current_value = '{}'.format(parent.GetLabel())
		
		if current_value == '...':
			current_value = ''
		
		self.text.SetValue(current_value)
		
		self.Bind(wx.EVT_TEXT_ENTER, self.on_text_enter)
		self.text.Bind(wx.EVT_KILL_FOCUS, self.on_focus_lost)
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)

		self.Move((parent.GetScreenPosition()))
		self.Show()


	def on_text_enter(self, event):
		self.parent.SetFocus()
		#this will trigger the frame to close


	def on_focus_lost(self, event):
		self.Close()
		event.Skip()


	def on_close_frame(self, event):
		#parse our the table name and field from the control name
		table = '.'.join(self.parent.GetName().split(':')[1].split('.')[0:2])
		table_id = wx.GetTopLevelParent(self.parent).id
		
		field = self.parent.GetName().split(':')[1].split('.')[2]
		new_value = self.text.GetValue()
		
		db.update_order(table, table_id, field, new_value)

		parent_frame = wx.GetTopLevelParent(self.parent)
		parent_frame.Freeze()
		parent_frame.reset_all()
		parent_frame.populate_all()
		parent_frame.Thaw()
		
		self.Destroy()


