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
		
		#for convenience, populate today's date when user focuses on a release field
		ctrl(self, 'text:orders.target_dates.actual_ae_release').Bind(wx.EVT_SET_FOCUS, self.on_focus_insert_date)
		ctrl(self, 'text:orders.target_dates.actual_de_release').Bind(wx.EVT_SET_FOCUS, self.on_focus_insert_date)
		ctrl(self, 'text:orders.target_dates.actual_mmg_release').Bind(wx.EVT_SET_FOCUS, self.on_focus_insert_date)

		#misc
		self.SetTitle('Item ID {}'.format(self.id))
		self.SetSize((976, 690))
		self.Center()
		
		#self.init_details_panel()
		self.init_changes_tab()
		
		self.populate_all()


		self.Show()

	
	def on_focus_insert_date(self, event):
		#for convenience, populate today's date when user focuses on a release field
		text_ctrl = event.GetEventObject()
		
		if text_ctrl.GetValue() == '':
			text_ctrl.SetValue(dt.date.today().strftime('%m/%d/%Y'))
	

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
		self.reset_labor_hours_tab()
		self.reset_changes_tab()
		
	def populate_all(self):
		self.populate_other_items_panel()
		self.populate_details_panel()
		self.populate_target_dates_tab()
		self.populate_labor_hours_tab()
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



	def reset_labor_hours_tab(self):
		ctrl(self, 'text:orders.labor_hours.applications_engineering').SetValue('')
		ctrl(self, 'text:orders.labor_hours.mechanical_engineering').SetValue('')
		ctrl(self, 'text:orders.labor_hours.electrical_engineering').SetValue('')
		ctrl(self, 'text:orders.labor_hours.structural_engineering').SetValue('')

		ctrl(self, 'text:orders.labor_hours.welding').SetValue('')
		ctrl(self, 'text:orders.labor_hours.painting').SetValue('')
		ctrl(self, 'text:orders.labor_hours.base_assembly').SetValue('')
		ctrl(self, 'text:orders.labor_hours.tube_fab_header').SetValue('')
		ctrl(self, 'text:orders.labor_hours.tube_fab').SetValue('')
		ctrl(self, 'text:orders.labor_hours.brazing').SetValue('')
		ctrl(self, 'text:orders.labor_hours.box_wire').SetValue('')
		ctrl(self, 'text:orders.labor_hours.hookup').SetValue('')
		ctrl(self, 'text:orders.labor_hours.testing').SetValue('')
		ctrl(self, 'text:orders.labor_hours.finishing').SetValue('')
		ctrl(self, 'text:orders.labor_hours.ship_loose').SetValue('')
		ctrl(self, 'text:orders.labor_hours.assembly').SetValue('')
		ctrl(self, 'text:orders.labor_hours.sheet_metal').SetValue('')



	def populate_labor_hours_tab(self):
		record = db.query('''
			SELECT
				applications_engineering,
				mechanical_engineering,
				electrical_engineering,
				structural_engineering,

				welding,
				painting,
				base_assembly,
				tube_fab_header,
				tube_fab,
				brazing,
				box_wire,
				hookup,
				testing,
				finishing,
				ship_loose,
				assembly,
				sheet_metal,
				
				(COALESCE(welding, 0) + COALESCE(painting, 0) + COALESCE(base_assembly, 0) +
					COALESCE(tube_fab_header, 0) + COALESCE(tube_fab, 0) + COALESCE(brazing, 0) +
					COALESCE(box_wire, 0) + COALESCE(hookup, 0) + COALESCE(testing, 0) +
					COALESCE(finishing, 0) + COALESCE(ship_loose, 0) + COALESCE(assembly, 0) +
					COALESCE(sheet_metal, 0)) AS hours_standard
			FROM
				orders.labor_hours
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
				
			else:
				field = '{}'.format(field)
				
			formatted_record.append(field)

		applications_engineering, mechanical_engineering, electrical_engineering, structural_engineering, \
		welding, painting, base_assembly, tube_fab_header, tube_fab, brazing, box_wire, hookup, testing, \
		finishing, ship_loose, assembly, sheet_metal, hours_standard = formatted_record

		ctrl(self, 'text:orders.labor_hours.applications_engineering').SetValue(applications_engineering)
		ctrl(self, 'text:orders.labor_hours.mechanical_engineering').SetValue(mechanical_engineering)
		ctrl(self, 'text:orders.labor_hours.electrical_engineering').SetValue(electrical_engineering)
		ctrl(self, 'text:orders.labor_hours.structural_engineering').SetValue(structural_engineering)

		ctrl(self, 'text:orders.labor_hours.welding').SetValue(welding)
		ctrl(self, 'text:orders.labor_hours.painting').SetValue(painting)
		ctrl(self, 'text:orders.labor_hours.base_assembly').SetValue(base_assembly)
		ctrl(self, 'text:orders.labor_hours.tube_fab_header').SetValue(tube_fab_header)
		ctrl(self, 'text:orders.labor_hours.tube_fab').SetValue(tube_fab)
		ctrl(self, 'text:orders.labor_hours.brazing').SetValue(brazing)
		ctrl(self, 'text:orders.labor_hours.box_wire').SetValue(box_wire)
		ctrl(self, 'text:orders.labor_hours.hookup').SetValue(hookup)
		ctrl(self, 'text:orders.labor_hours.testing').SetValue(testing)
		ctrl(self, 'text:orders.labor_hours.finishing').SetValue(finishing)
		ctrl(self, 'text:orders.labor_hours.ship_loose').SetValue(ship_loose)
		ctrl(self, 'text:orders.labor_hours.assembly').SetValue(assembly)
		ctrl(self, 'text:orders.labor_hours.sheet_metal').SetValue(sheet_metal)
		
		ctrl(self, 'label:standard_hours').SetLabel('{}'.format(hours_standard))



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
				address,
				
				bpcs_sales_order,
				bpcs_line_up,
				bpcs_item,
				bpcs_family

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
		sold_to_name, sold_to_number, ship_to_name, ship_to_number, country, state, city, zip_code, address, \
		bpcs_sales_order, bpcs_line_up, bpcs_item, bpcs_family = formatted_record
		
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
		
		#format the BPCS Sales Order and Line Up together
		bpcs_sales_order_and_line_up = '{}-{}'.format(bpcs_sales_order, bpcs_line_up)
		if bpcs_sales_order_and_line_up == '...-...':
			bpcs_sales_order_and_line_up = '...'

		ctrl(self, 'label:bpcs_item').SetLabel(bpcs_item)
		ctrl(self, 'label:bpcs_sales_order').SetLabel(bpcs_sales_order_and_line_up)
		ctrl(self, 'label:bpcs_family').SetLabel(bpcs_family)


	def init_changes_tab(self):
		column_names = ['Id', 'Table', 'Field', 'Previous Value', 'New Value', 'Who Changed', 'When Changed']

		for index, column_name in enumerate(column_names):
			ctrl(self, 'list:changes').InsertColumn(index, column_name)

	def reset_changes_tab(self):
		ctrl(self, 'list:changes').DeleteAllItems()

	def populate_changes_tab(self):
		list_ctrl = ctrl(self, 'list:changes')
		list_ctrl.Freeze()
		
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
				id DESC
			'''.format(self.id))
		
		#insert records into list
		for index, record in enumerate(records):
			id, table_name, field, previous_value, new_value, who_changed, when_changed = record
			
			#convert values to formatted datetimes if possible
			try:
				if len(previous_value) == 19 and previous_value.count(':') == 2 and previous_value.count('-') == 2:
					if '00:00:00' in previous_value:
						previous_value = dt.datetime.strptime(previous_value, "%Y-%m-%d %H:%M:%S").strftime('%m/%d/%Y')
					else:
						previous_value = dt.datetime.strptime(previous_value, "%Y-%m-%d %H:%M:%S").strftime('%m/%d/%y %I:%M %p')

				if len(new_value) == 19 and new_value.count(':') == 2 and new_value.count('-') == 2:
					if '00:00:00' in new_value:
						new_value = dt.datetime.strptime(new_value, "%Y-%m-%d %H:%M:%S").strftime('%m/%d/%Y')
					else:
						new_value = dt.datetime.strptime(new_value, "%Y-%m-%d %H:%M:%S").strftime('%m/%d/%y %I:%M %p')
			except:
				pass
				
			#blank out null or none values
			if previous_value == None or previous_value == 'NULL':
				previous_value = ''

			if new_value == None or new_value == 'NULL':
				new_value = ''

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
		
		list_ctrl.Thaw()


	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()
