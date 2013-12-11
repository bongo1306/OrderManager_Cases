#!/usr/bin/env python
# -*- coding: utf8 -*-

import pyodbc
import sys
import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import datetime as dt
import TweakedGrid
import General as gn
import Database as db


class SearchTab(object):
	def init_search_tab(self):
		#this will hold the fleeting search criteria tables
		self.table_search_criteria = None
		
		#Bindings
		self.Bind(wx.EVT_CHOICE, self.on_choice_table, id=xrc.XRCID('choice:which_table'))
		self.Bind(wx.EVT_CHECKBOX, self.on_choice_table, id=xrc.XRCID('checkbox:display_alphabetically'))
		
		#tables or views the user can search in
		tables = ('orders.root', 'orders.view_systems', 'dbo.orders', 'dbo.view_orders_old')
		
		ctrl(self, 'choice:which_table').AppendItems(tables)
		ctrl(self, 'choice:which_table').SetStringSelection('orders.view_systems')
		self.on_choice_table()
		


	def on_choice_table(self, event=None):
		table_panel = ctrl(self, 'panel:search_criteria')
		
		#remove the table if there is already one there
		if self.table_search_criteria != None:
			self.table_search_criteria.Destroy()
		
		self.table_search_criteria = TweakedGrid.TweakedGrid(table_panel)
		
		table_to_search = ctrl(self, 'choice:which_table').GetStringSelection()

		columns = list(db.get_table_column_names(table_to_search, presentable=False))
		
		#manually specify columns for some tables...
		try:
			columns = list(custom_search_table_columns[table_to_search])
		except:
			pass
		
		for column_index, column in enumerate(columns):
			if '_spacer_' in column:
				column = ''
			columns[column_index] = '{}'.format(column)


		#sort columns alphabetically if user wants
		if ctrl(self, 'checkbox:display_alphabetically').Value == True:
			#columns = list(set(columns))#.remove('')
			columns = filter(None, columns)
			columns.sort()

		self.table_search_criteria.CreateGrid(len(columns), 2)
		self.table_search_criteria.SetRowLabelSize(0)
		self.table_search_criteria.SetColLabelValue(0, 'Field')
		self.table_search_criteria.SetColLabelValue(1, 'Criteria')
		
		for column_index, column in enumerate(columns):
			if column != '':
				self.table_search_criteria.SetCellValue(column_index, 0, column)
				self.table_search_criteria.SetCellValue(column_index, 1, '')
			else:
				self.table_search_criteria.SetReadOnly(column_index, 1)
				#self.table_search_criteria.SetCellValue(column_index, 0, column)
				#self.table_search_criteria.SetCellValue(column_index, 1, 'Fields from {}:'.format(column[1:]))
		
		#self.table_search_criteria.SetCellValue(0, 0,' (click to select document) ')
		self.table_search_criteria.AutoSize()
		
		self.table_search_criteria.EnableDragRowSize(False)
		
		
		self.table_search_criteria.Bind(wx.EVT_SIZE, self.on_size_criteria_table)
		self.table_search_criteria.Bind(wx.grid.EVT_GRID_CELL_CHANGE, self.on_change_grid_cell)
		#self.table_search_criteria.Bind(wx.EVT_CHAR, on_change_grid_cell)
		
		for row in range(len(columns)):
			self.table_search_criteria.SetReadOnly(row, 0)
			#self.table_search_criteria.SetCellAlignment(row, 0, wx.ALIGN_RIGHT, wx.ALIGN_CENTRE)
		
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(self.table_search_criteria, 1, wx.EXPAND)
		table_panel.SetSizer(sizer)
		
		table_panel.Layout()


	def on_size_criteria_table(self, event):
		table = event.GetEventObject()
		table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0) - wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X))

		event.Skip()


	def on_change_grid_cell(self, event):
		#ctrl(self, 'text:sql_query').SetValue(self.generate_sql_query())
		print 'HELLO!!!!!!!!!!'
		event.Skip()

