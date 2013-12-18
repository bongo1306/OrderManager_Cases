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
		self.Bind(wx.EVT_BUTTON, self.on_click_goto_next_item, id=xrc.XRCID('button:next_item'))
		self.Bind(wx.EVT_BUTTON, self.on_click_goto_previous_item, id=xrc.XRCID('button:previous_item'))
		
		#specify what tables the various controls should try to update
		#ctrl(self, 'text:requested_mmg_release').table = 'orders.target_dates'
		#ctrl(self, 'text:requested_mmg_release').table = 'orders.target_dates'

		#misc
		self.SetTitle('Item ID {}'.format(self.id))
		self.SetSize((976, 690))
		self.Center()
		
		#self.init_details_panel()
		self.init_changes_tab()
		
		self.populate_all()


		self.Show()

	
	def on_focus_lost(self, event):
		print 'lost focus!', event
	

	def on_click_goto_previous_item(self, event):
		previous_id = db.query('''
			SELECT TOP 1
				id
			FROM
				orders.root
			WHERE
				sales_order=(SELECT TOP 1 sales_order FROM orders.root WHERE id={}) AND
				CAST(item AS INT) < (SELECT TOP 1 CAST(item AS INT) FROM orders.root WHERE id={})
			ORDER BY
				CAST(item AS INT) DESC
			'''.format(self.id, self.id))[0]

		self.Freeze()
		self.reset_all()

		self.id = previous_id
		self.populate_all()
		self.Thaw()
		
	def on_click_goto_next_item(self, event):
		next_id = db.query('''
			SELECT
				id
			FROM
				orders.root
			WHERE
				sales_order=(SELECT TOP 1 sales_order FROM orders.root WHERE id={}) AND
				CAST(item AS INT) > (SELECT TOP 1 CAST(item AS INT) FROM orders.root WHERE id={})
			ORDER BY
				CAST(item AS INT) ASC
			'''.format(self.id, self.id))[0]

		self.Freeze()
		self.reset_all()
		
		self.id = next_id
		self.populate_all()
		self.Thaw()
		

	def reset_all(self):
		self.reset_other_items_panel()
		self.reset_details_panel()
		self.reset_target_dates_tab()
		self.reset_changes_tab()
		
	def populate_all(self):
		self.populate_other_items_panel()
		self.populate_details_panel()
		self.populate_target_dates_tab()
		self.populate_changes_tab()

		ctrl(self, 'panel:main').Layout()


	def reset_target_dates_tab(self):
		ctrl(self, 'text:orders.target_dates.requested_ae_release').SetValue('')
		ctrl(self, 'text:orders.target_dates.planned_ae_release').SetValue('')
		ctrl(self, 'text:orders.target_dates.suggested_ae_start').SetValue('')
		ctrl(self, 'text:orders.target_dates.actual_ae_release').SetValue('')
		ctrl(self, 'checkbox:orders.target_dates.planned_ae_release_locked').SetValue(False)

		ctrl(self, 'text:orders.target_dates.requested_de_release').SetValue('')
		ctrl(self, 'text:orders.target_dates.planned_de_release').SetValue('')
		ctrl(self, 'text:orders.target_dates.suggested_de_start').SetValue('')
		ctrl(self, 'text:orders.target_dates.actual_de_release').SetValue('')
		ctrl(self, 'checkbox:orders.target_dates.planned_de_release_locked').SetValue(False)

		ctrl(self, 'text:orders.target_dates.requested_mmg_release').SetValue('')
		ctrl(self, 'text:orders.target_dates.planned_mmg_release').SetValue('')
		ctrl(self, 'text:orders.target_dates.suggested_mmg_start').SetValue('')
		ctrl(self, 'text:orders.target_dates.actual_mmg_release').SetValue('')
		ctrl(self, 'checkbox:orders.target_dates.planned_mmg_release_locked').SetValue(False)


	def populate_target_dates_tab(self):
		record = db.query('''
			SELECT
				requested_ae_release,
				planned_ae_release,
				planned_ae_release_locked,
				suggested_ae_start,
				actual_ae_release,
				
				requested_de_release,
				planned_de_release,
				planned_de_release_locked,
				suggested_de_start,
				actual_de_release,
				
				requested_mmg_release,
				planned_mmg_release,
				planned_mmg_release_locked,
				suggested_mmg_start,
				actual_mmg_release
			FROM
				orders.target_dates
			WHERE
				id={}
			'''.format(self.id))
			
		if not record:
			return
			
		#format all fields as strings
		formatted_record = []
		for field in record[0]:
			if field == None:
				field = ''
				
			elif isinstance(field, dt.datetime):
				field = field.strftime('%m/%d/%Y')
				
			else:
				pass
				
			formatted_record.append(field)

		requested_ae_release, planned_ae_release, planned_ae_release_locked, suggested_ae_start, actual_ae_release, \
		requested_de_release, planned_de_release, planned_de_release_locked, suggested_de_start, actual_de_release, \
		requested_mmg_release, planned_mmg_release, planned_mmg_release_locked, suggested_mmg_start, actual_mmg_release = formatted_record

		ctrl(self, 'text:orders.target_dates.requested_ae_release').SetValue(requested_ae_release)
		ctrl(self, 'text:orders.target_dates.planned_ae_release').SetValue(planned_ae_release)
		ctrl(self, 'text:orders.target_dates.suggested_ae_start').SetValue(suggested_ae_start)
		ctrl(self, 'text:orders.target_dates.actual_ae_release').SetValue(actual_ae_release)
		ctrl(self, 'checkbox:orders.target_dates.planned_ae_release_locked').SetValue(planned_ae_release_locked)

		ctrl(self, 'text:orders.target_dates.requested_de_release').SetValue(requested_de_release)
		ctrl(self, 'text:orders.target_dates.planned_de_release').SetValue(planned_de_release)
		ctrl(self, 'text:orders.target_dates.suggested_de_start').SetValue(suggested_de_start)
		ctrl(self, 'text:orders.target_dates.actual_de_release').SetValue(actual_de_release)
		ctrl(self, 'checkbox:orders.target_dates.planned_de_release_locked').SetValue(planned_de_release_locked)

		ctrl(self, 'text:orders.target_dates.requested_mmg_release').SetValue(requested_mmg_release)
		ctrl(self, 'text:orders.target_dates.planned_mmg_release').SetValue(planned_mmg_release)
		ctrl(self, 'text:orders.target_dates.suggested_mmg_start').SetValue(suggested_mmg_start)
		ctrl(self, 'text:orders.target_dates.actual_mmg_release').SetValue(actual_mmg_release)
		ctrl(self, 'checkbox:orders.target_dates.planned_mmg_release_locked').SetValue(planned_mmg_release_locked)


	def reset_other_items_panel(self):
		ctrl(self, 'label:other_items').SetLabel('Item X of X')
		ctrl(self, 'button:previous_item').Enable()
		ctrl(self, 'button:next_item').Enable()

	def populate_other_items_panel(self):
		all_items = db.query('''
			SELECT
				id, CAST(item AS INT)
			FROM
				orders.root
			WHERE
				sales_order=(SELECT TOP 1 sales_order FROM orders.root WHERE id={})
			ORDER BY
				CAST(item AS INT) DESC
			'''.format(self.id))
		
		#determine the current item given id without querying the database again :)
		current_item = None
		for item_data in all_items:
			if item_data[0] == self.id:
				current_item = item_data[1]
				
		try:
			max_item = all_items[0][1]
		except:
			max_item = None

		try:
			min_item = all_items[-1][1]
		except:
			min_item = None

		ctrl(self, 'label:other_items').SetLabel('Item {} of {}'.format(current_item, max_item))
		self.SetTitle('Item ID {}'.format(self.id))
		
		#disable previous or next buttons if there are no more items
		if current_item <= min_item:
			ctrl(self, 'button:previous_item').Disable()
		
		if current_item >= max_item:
			ctrl(self, 'button:next_item').Disable()




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
				description,
				
				sold_to_name,
				sold_to_number,
				ship_to_name,
				ship_to_number,
				country,
				state,
				city,
				zip_code,
				address

			FROM
				orders.root
			WHERE
				id={}
			'''.format(self.id))
		
		if not record:
			return
			
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

		filemaker_quote, sales_order, item, production_order, material, hierarchy, model, description, \
		sold_to_name, sold_to_number, ship_to_name, ship_to_number, country, state, city, zip_code, address = formatted_record
		
		ctrl(self, 'label:quote').SetLabel(filemaker_quote)
		ctrl(self, 'label:sales_order').SetLabel(sales_order)
		ctrl(self, 'label:item').SetLabel(item)
		ctrl(self, 'label:production_order').SetLabel(production_order)
		ctrl(self, 'label:material').SetLabel(material)
		ctrl(self, 'label:hierarchy').SetLabel(hierarchy)
		ctrl(self, 'label:model').SetLabel(model)
		ctrl(self, 'label:description').SetLabel(description)

		ctrl(self, 'label:sold_to').SetLabel('{} ({})'.format(sold_to_name, sold_to_number))
		ctrl(self, 'label:ship_to').SetLabel('{} ({})'.format(sold_to_name, sold_to_number))
		ctrl(self, 'label:address').SetLabel(address)
		ctrl(self, 'label:city_state').SetLabel('{}, {} ({}) {}'.format(city, state, country, zip_code))
		


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
