#!/usr/bin/env python
# -*- coding: utf8 -*-

import pyodbc
import sys
import wx

import datetime as dt
import TweakedGrid
import General as gn


class SearchTab(object):
	def init_search_tab(self):
		#this will hold the fleeting search criteria tables
		self.table_search_criteria = None
		
		#Bindings
		self.Bind(wx.EVT_CHOICE, self.on_choice_table, id=xrc.XRCID('choice:which_table'))
		
		#tables or views the user can search in
		tables = ('orders.root', 'orders.view_systems', 'dbo.orders', 'dbo.view_orders_old')
		
		ctrl(self, 'choice:which_table').AppendItems(tables)
		ctrl(self, 'choice:which_table').SetStringSelection('orders.view_systems')
		


	def on_choice_table(self):
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
			columns[column_index] = '{}.{}'.format(table_to_search, column)

		for joining_table in self.joining_tables:
			if table_to_search == joining_table[0]:
				extend_list = ['']
				#extend_list.append('-{}'.format(joining_table[1]))
				
				#manually specify columns for some tables...
				try:
					extended_columns = list(custom_search_table_columns[joining_table[1]])
				except:
					extended_columns = list(db.get_table_column_names(joining_table[1], presentable=False))
				
				extend_list.extend(extended_columns)
				extend_list.remove(joining_table[3])
				
				for column_index, column in enumerate(extend_list):
					if column_index > 0:
						extend_list[column_index] = '{}.{}'.format(joining_table[1], column)
						
				columns.extend(extend_list)




		#sort columns alphabeticaly <--(lol spelling)
		#columns.sort()


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
		
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(self.table_search_criteria, 1, wx.EXPAND)
		table_panel.SetSizer(sizer)
		
		table_panel.Layout()


	def on_size_criteria_table(self, event):
		table = event.GetEventObject()
		table.SetColSize(1, table.GetSize()[0] - table.GetColSize(0) - wx.SystemSettings.GetMetric(wx.SYS_VSCROLL_X))

		event.Skip()


	def on_change_grid_cell(self, event):
		ctrl(self, 'text:sql_query').SetValue(self.generate_sql_query())
		event.Skip()

