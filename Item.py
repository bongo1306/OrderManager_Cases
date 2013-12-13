#!/usr/bin/env python
# -*- coding: utf8 -*-

import sys

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

		self.init_details_panel()
		self.populate_details_panel()

		self.init_changes_tab()
		self.populate_changes_tab()

		#misc
		self.SetTitle('Item ID {}'.format(self.id))
		self.SetSize((800, 600))
		self.Center()


		self.Show()



	def init_details_panel(self):
		pass

	def reset_details_panel(self):
		ctrl(self, 'label:quote').SetLabel('...')
		ctrl(self, 'label:sales_order').SetLabel('...')

	def populate_details_panel(self):
		record = db.query('''
			SELECT
				filemaker_quote,
				sales_order,
				item,
				production_order,
				material,
				hierarchy,
				model,
				description
			FROM
				orders.root
			WHERE
				id={}
			'''.format(self.id))
			
		#format all fields as strings
		formatted_record = []
		for field in record[0]:
			if field == None:
				field = '...'
				
			elif isinstance(field, dt.datetime):
				field = field.strftime('%m/%d/%y')
				
			else:
				field = str(field)
				
			formatted_record.append(field)

		filemaker_quote, sales_order, item, production_order, material, hierarchy, model, description = formatted_record
		
		ctrl(self, 'label:quote').SetLabel(filemaker_quote)
		ctrl(self, 'label:sales_order').SetLabel(sales_order)
		ctrl(self, 'label:item').SetLabel(item)
		ctrl(self, 'label:production_order').SetLabel(production_order)
		ctrl(self, 'label:material').SetLabel(material)
		ctrl(self, 'label:hierarchy').SetLabel(hierarchy)
		ctrl(self, 'label:model').SetLabel(model)
		ctrl(self, 'label:description').SetLabel(description)



	def init_changes_tab(self):
		column_names = ['Id', 'Table', 'Field', 'Previous Value', 'New Value', 'Who Changed', 'When Changed']

		for index, column_name in enumerate(column_names):
			ctrl(self, 'list:changes').InsertColumn(index, column_name)

	def reset_changes_tab(self):
		ctrl(self, 'list:changes').DeleteAllItems()

	def populate_changes_tab(self):
		list_ctrl = ctrl(self, 'list:changes')
		
		records = db.query('''
			SELECT
				id,
				table_name,
				field,
				previous_value,
				new_value,
				who_changed,
				when_changed
			FROM
				orders.changes
			WHERE
				table_id={}
			ORDER BY
				when_changed DESC
			'''.format(self.id))
		
		#insert records into list
		for index, record in enumerate(records):
			id, table_name, field, previous_value, new_value, who_changed, when_changed = record
			
			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(table_name))
			list_ctrl.SetStringItem(index, 2, '{}'.format(field))
			list_ctrl.SetStringItem(index, 3, '{}'.format(previous_value))
			list_ctrl.SetStringItem(index, 4, '{}'.format(new_value))
			list_ctrl.SetStringItem(index, 5, '{}'.format(who_changed))
			list_ctrl.SetStringItem(index, 6, '{}'.format(when_changed.strftime('%m/%d/%y %I:%M %p')))

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)
			
			#cap column width at max 400
			if list_ctrl.GetColumnWidth(index) > 400:
				list_ctrl.SetColumnWidth(index, 400)
		
		#hide id column
		list_ctrl.SetColumnWidth(0, 0)


	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()
