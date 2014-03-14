#!/usr/bin/env python
# -*- coding: utf8 -*-

import sys
import os
import subprocess

import SnapshotPrinter		#for printing of wx window

import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import datetime as dt
import General as gn
import Database as db

import TweakedGrid


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

		self.Bind(wx.EVT_BUTTON, self.on_click_open_order_folder, id=xrc.XRCID('button:order_folder'))

		self.Bind(wx.EVT_BUTTON, self.on_click_goto_next_item, id=xrc.XRCID('button:next_item'))
		self.Bind(wx.EVT_BUTTON, self.on_click_goto_previous_item, id=xrc.XRCID('button:previous_item'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_print_window, id=xrc.XRCID('button:print'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_applications_engineer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_design_engineer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_mechanical_engineer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_electrical_engineer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_structural_engineer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_mechanical_cad_designer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_electrical_cad_designer'))
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time_structural_cad_designer'))

		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log_time'))

		#for convenience, populate today's date when user focuses on a release field
		ctrl(self, 'text:orders.target_dates.actual_ae_release').Bind(wx.EVT_SET_FOCUS, self.on_focus_insert_date)
		ctrl(self, 'text:orders.target_dates.actual_de_release').Bind(wx.EVT_SET_FOCUS, self.on_focus_insert_date)
		ctrl(self, 'text:orders.target_dates.actual_de_release').Bind(wx.EVT_KILL_FOCUS, self.on_kill_focus_check_if_sent_to_mmg)
		ctrl(self, 'text:orders.target_dates.actual_mmg_release').Bind(wx.EVT_SET_FOCUS, self.on_focus_insert_date)

		self.applicable_labor_hours_per_material = {
			'CDA': ('welding', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'CPP': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'CSS': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'CTL': ('box_wire', 'testing', 'finishing'), 
			'CVS': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'DBV': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'DSP': ('welding', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'DSS': ('welding', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'DSSIIX': ('welding', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'DXVS': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'FAH': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'FAV': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'FAX': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'HPM': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'HVS': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'MISC': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'NH2': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'NHS': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'NV2': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'NX2': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'OHD': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'OHN': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'OHS': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'OHW': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'ONH': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'ONV': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'PSM': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'RCA': ('welding', 'painting', 'base_assembly', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'RHD': ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 'brazing', 'box_wire', 'hookup', 'testing', 'finishing'), 
			'SHIP_LOOSE': ('ship_loose', ), 
			'SPARTCOLS': ('finishing', ), 
			'WEE': ('welding', 'painting', 'tube_fab', 'brazing', 'box_wire', 'testing', 'finishing', 'assembly'), 
			'WEH': ('welding', 'painting', 'tube_fab', 'brazing', 'box_wire', 'testing', 'finishing', 'assembly'), 
			'WEM': ('welding', 'painting', 'tube_fab', 'brazing', 'box_wire', 'testing', 'finishing', 'assembly'), 
		}
		
		#misc
		self.SetTitle('Item ID {}'.format(self.id))
		self.SetSize((976, 690))
		self.Center()
		
		self.init_details_panel()
		self.init_responsibilities_tab()
		self.init_changes_tab()
		self.init_time_logs_tab()
		self.init_tabulated_data_tab()
		
		self.populate_all()


		self.Show()


	def on_click_print_window(self, event):
		SnapshotPrinter.print_window(self)
		wx.CallAfter(self.Raise)

	
	def on_focus_insert_date(self, event):
		#for convenience, populate today's date when user focuses on a release field
		text_ctrl = event.GetEventObject()
		
		if text_ctrl.GetValue() == '':
			text_ctrl.SetValue(dt.date.today().strftime('%m/%d/%Y'))
		
		event.Skip()


	def on_kill_focus_check_if_sent_to_mmg(self, event):
		production_order = db.query("SELECT production_order FROM orders.root WHERE id={}".format(self.id))[0]
		
		result = db.query("SELECT TOP 1 production_order FROM dbo.mmg_uploads WHERE production_order='{}'".format(production_order))
		
		if not result and event.GetEventObject().GetValue() != '':
			wx.MessageBox("An entry in dbo.mmg_uploads for this production order was not found.\nThis probably means the BOM for this order still needs to be processed and sent to MMG.\n\nThe MeatGrinder will now open to encourage you to send {} to MMG.".format(production_order), 'Notice', wx.OK | wx.ICON_INFORMATION)
			os.startfile(r"\\kw_engineering\sharepoint$\Everyone\Management Software\eng2mmgBOMupload\eng2mmgBOMupload.xlsm")
		
		event.Skip()

	
	def on_click_open_order_folder(self, event):
		event.GetEventObject().Disable()
		
		sales_orders = db.query('''
			SELECT TOP 1
				sales_order,
				bpcs_sales_order
			FROM
				orders.root
			WHERE
				id={}
			'''.format(self.id))
		
		if sales_orders:
			sap_so, bpcs_so = sales_orders[0]
		else:
			sap_so, bpcs_so = (None, None)
		
		a_folder_was_found = False
		
		if sap_so:
			sap_order_folder_path = self.find_sap_order_folder_path(sap_so)
			
			if sap_order_folder_path:
				subprocess.Popen('explorer "{}"'.format(sap_order_folder_path))
				a_folder_was_found = True

		if bpcs_so:
			bpcs_order_folder_path = self.find_bpcs_order_folder_path(bpcs_so)
			
			if bpcs_order_folder_path:
				subprocess.Popen('explorer "{}"'.format(bpcs_order_folder_path))
				a_folder_was_found = True

		if a_folder_was_found == False:
			wx.MessageBox("A folder for this order was not automatically found.\nMaybe it doesn't exist yet or is named incorrectly.", 'Notice', wx.OK | wx.ICON_INFORMATION)

		event.GetEventObject().Enable()


	def find_bpcs_order_folder_path(self, bpcs_so):
		starting_path = r"\\kw_engineering\eng_res\Design_Eng\Orders\Orders_20{}".format(bpcs_so[1:3])

		#plow through three directories deep looking for a folder named that bpcs sales order
		for x in os.listdir(starting_path):
			starting_path_x = os.path.join(starting_path, x)
			
			if os.path.isdir(starting_path_x):
				for y in os.listdir(starting_path_x):
					starting_path_x_y = os.path.join(starting_path_x, y)

					if os.path.isdir(starting_path_x_y):
						for z in os.listdir(starting_path_x_y):
							starting_path_x_y_z = os.path.join(starting_path_x_y, z)
							
							if os.path.isdir(starting_path_x_y_z):
								if z == bpcs_so:
									return starting_path_x_y_z
		
		return None


	def find_sap_order_folder_path(self, sap_so):
		starting_path = r"\\kw_engineering\eng_res\Design_Eng\Orders\SAP_ORDERS_COLS"

		#plow through two directories deep looking for a folder named that sap sales order
		for x in os.listdir(starting_path):
			starting_path_x = os.path.join(starting_path, x)
			
			if os.path.isdir(starting_path_x):
				for y in os.listdir(starting_path_x):
					starting_path_x_y = os.path.join(starting_path_x, y)
					
					if os.path.isdir(starting_path_x_y):
						if y == sap_so:
							return starting_path_x_y
		
		return None


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
		self.reset_responsibilities_tab()
		self.reset_target_dates_tab()
		self.reset_labor_hours_tab()
		self.reset_changes_tab()
		self.reset_time_logs_tab()
		self.reset_misc_tab()
		self.reset_financials_tab()
		self.reset_tabulated_data_tab()
		

	def populate_all(self):
		self.populate_other_items_panel()
		self.populate_details_panel()
		self.populate_responsibilities_tab()
		self.populate_target_dates_tab()
		self.populate_labor_hours_tab()
		self.populate_changes_tab()
		self.populate_time_logs_tab()
		self.populate_misc_tab()
		self.populate_financials_tab()
		self.populate_tabulated_data_tab()

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

		ctrl(self, 'label:date_created_on').SetLabel('...')
		ctrl(self, 'label:date_basic_start').SetLabel('...')
		ctrl(self, 'label:date_basic_finish').SetLabel('...')
		ctrl(self, 'label:date_actual_finish').SetLabel('...')
		ctrl(self, 'label:date_request_delivered').SetLabel('...')
		ctrl(self, 'label:date_shipped').SetLabel('...')
		
		ctrl(self, 'checkbox:orders.misc.pending_ecms').SetValue(False)


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
			
		if record:
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

			if planned_ae_release_locked == '':
				planned_ae_release_locked = False

			if planned_de_release_locked == '':
				planned_de_release_locked = False

			if planned_mmg_release_locked == '':
				planned_mmg_release_locked = False

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


		record = db.query('''
			SELECT
				date_created_on,
				date_basic_start,
				date_basic_finish,
				date_actual_finish,
				date_request_delivered,
				date_shipped
			FROM
				orders.root
			WHERE
				id={}
			'''.format(self.id))
			
		if record:
			#format all fields as strings
			formatted_record = []
			for field in record[0]:
				if field == None:
					field = '...'
					
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%Y')
					
				else:
					pass
					
				formatted_record.append(field)

			date_created_on, date_basic_start, date_basic_finish, date_actual_finish, date_request_delivered, date_shipped = formatted_record

			ctrl(self, 'label:date_created_on').SetLabel(date_created_on)
			ctrl(self, 'label:date_basic_start').SetLabel(date_basic_start)
			ctrl(self, 'label:date_basic_finish').SetLabel(date_basic_finish)
			ctrl(self, 'label:date_actual_finish').SetLabel(date_actual_finish)
			ctrl(self, 'label:date_request_delivered').SetLabel(date_request_delivered)
			ctrl(self, 'label:date_shipped').SetLabel(date_shipped)


		record = db.query('''
			SELECT
				pending_ecms
			FROM
				orders.misc
			WHERE
				id={}
			'''.format(self.id))
	
		if record:
			pending_ecms = record[0]
			ctrl(self, 'checkbox:orders.misc.pending_ecms').SetValue(pending_ecms)


	def reset_misc_tab(self):
		ctrl(self, 'combo:orders.misc.refrigerant').SetValue('')
		ctrl(self, 'combo:orders.misc.compressor_manufacturer').SetValue('')
		ctrl(self, 'combo:orders.misc.compressor_type').SetValue('')
		ctrl(self, 'text:orders.misc.compressor_quantity').SetValue('')
		ctrl(self, 'text:orders.misc.circuit_quantity').SetValue('')
		ctrl(self, 'combo:orders.misc.controller').SetValue('')
		ctrl(self, 'text:orders.misc.length').SetValue('')
		ctrl(self, 'text:orders.misc.width').SetValue('')
		ctrl(self, 'text:orders.misc.height').SetValue('')


	def populate_misc_tab(self):
		record = db.query('''
			SELECT
				refrigerant,
				compressor_manufacturer,
				compressor_type,
				compressor_quantity,
				circuit_quantity,
				controller,
				length,
				width,
				height
			FROM
				orders.misc
			WHERE
				id={}
			'''.format(self.id))

		if record:
			#format all fields as strings
			formatted_record = []
			for field in record[0]:
				if field == None:
					field = ''
				else:
					pass
					
				formatted_record.append(field)

			refrigerant, compressor_manufacturer, compressor_type, compressor_quantity, \
			circuit_quantity, controller, length, width, height = formatted_record

			ctrl(self, 'combo:orders.misc.refrigerant').SetValue(refrigerant)
			ctrl(self, 'combo:orders.misc.compressor_manufacturer').SetValue(compressor_manufacturer)
			ctrl(self, 'combo:orders.misc.compressor_type').SetValue(compressor_type)
			ctrl(self, 'text:orders.misc.compressor_quantity').SetValue(compressor_quantity)
			ctrl(self, 'text:orders.misc.circuit_quantity').SetValue(circuit_quantity)
			ctrl(self, 'combo:orders.misc.controller').SetValue(controller)
			ctrl(self, 'text:orders.misc.length').SetValue(length)
			ctrl(self, 'text:orders.misc.width').SetValue(width)
			ctrl(self, 'text:orders.misc.height').SetValue(height)


	def reset_financials_tab(self):
		ctrl(self, 'text:orders.financials.material_quote').SetValue('')
		ctrl(self, 'text:orders.financials.labor_quote').SetValue('')
		ctrl(self, 'text:orders.financials.overhead_quote').SetValue('')
		ctrl(self, 'text:orders.financials.cost_quote').SetValue('')
		ctrl(self, 'text:orders.financials.copper_surcharge_quote').SetValue('')
		ctrl(self, 'text:orders.financials.steel_surcharge_quote').SetValue('')
		ctrl(self, 'text:orders.financials.material_standard').SetValue('')
		ctrl(self, 'text:orders.financials.labor').SetValue('')
		ctrl(self, 'text:orders.financials.overhead').SetValue('')
		ctrl(self, 'text:orders.financials.cost_standard').SetValue('')
		ctrl(self, 'text:orders.financials.list_price').SetValue('')
		ctrl(self, 'text:orders.financials.discount_multiplier').SetValue('')
		ctrl(self, 'text:orders.financials.margin_standard').SetValue('')
		ctrl(self, 'text:orders.financials.gross_margin_standard').SetValue('')
		ctrl(self, 'text:orders.financials.net_sales').SetValue('')


	def populate_financials_tab(self):
		record = db.query('''
			SELECT
				material_quote, 
				labor_quote, 
				overhead_quote, 
				cost_quote, 
				copper_surcharge_quote, 
				steel_surcharge_quote, 
				material_standard, 
				labor, 
				overhead, 
				cost_standard, 
				list_price, 
				discount_multiplier, 
				margin_standard, 
				gross_margin_standard, 
				net_sales
			FROM
				orders.financials
			WHERE
				id={}
			'''.format(self.id))

		if record:
			#format all fields as strings
			formatted_record = []
			for field in record[0]:
				if field == None:
					field = ''
				else:
					field = '{}'.format(field)
					
				formatted_record.append(field)

			material_quote, labor_quote, overhead_quote, cost_quote, copper_surcharge_quote, steel_surcharge_quote, \
			material_standard, labor, overhead, cost_standard, list_price, discount_multiplier, \
			margin_standard, gross_margin_standard, net_sales = formatted_record

			ctrl(self, 'text:orders.financials.material_quote').SetValue(material_quote)
			ctrl(self, 'text:orders.financials.labor_quote').SetValue(labor_quote)
			ctrl(self, 'text:orders.financials.overhead_quote').SetValue(overhead_quote)
			ctrl(self, 'text:orders.financials.cost_quote').SetValue(cost_quote)
			ctrl(self, 'text:orders.financials.copper_surcharge_quote').SetValue(copper_surcharge_quote)
			ctrl(self, 'text:orders.financials.steel_surcharge_quote').SetValue(steel_surcharge_quote)
			ctrl(self, 'text:orders.financials.material_standard').SetValue(material_standard)
			ctrl(self, 'text:orders.financials.labor').SetValue(labor)
			ctrl(self, 'text:orders.financials.overhead').SetValue(overhead)
			ctrl(self, 'text:orders.financials.cost_standard').SetValue(cost_standard)
			ctrl(self, 'text:orders.financials.list_price').SetValue(list_price)
			ctrl(self, 'text:orders.financials.discount_multiplier').SetValue(discount_multiplier)
			ctrl(self, 'text:orders.financials.margin_standard').SetValue(margin_standard)
			ctrl(self, 'text:orders.financials.gross_margin_standard').SetValue(gross_margin_standard)
			ctrl(self, 'text:orders.financials.net_sales').SetValue(net_sales)


	def reset_labor_hours_tab(self):
		ctrl(self, 'text:orders.labor_hours.applications_engineering').SetValue('')
		ctrl(self, 'text:orders.labor_hours.mechanical_engineering').SetValue('')
		ctrl(self, 'text:orders.labor_hours.electrical_engineering').SetValue('')
		ctrl(self, 'text:orders.labor_hours.structural_engineering').SetValue('')

		text_ctrl = ctrl(self, 'text:orders.labor_hours.welding')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.painting')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.base_assembly')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.tube_fab_header')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.tube_fab')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.brazing')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.box_wire')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.hookup')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.testing')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.finishing')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.ship_loose')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.assembly')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)

		text_ctrl = ctrl(self, 'text:orders.labor_hours.sheet_metal')
		text_ctrl.SetValue('')
		text_ctrl.SetBackgroundColour(wx.NullColour)
		text_ctrl.SetForegroundColour(wx.NullColour)


		ctrl(self, 'label:standard_hours').SetLabel('...')



	def populate_labor_hours_tab(self):
		#darken the fields that do not apply to this material
		material = db.query("SELECT material FROM orders.root WHERE id={}".format(self.id))
		
		labor_hour_fields = ('welding', 'painting', 'base_assembly', 'tube_fab_header', 'tube_fab', 
			'brazing', 'box_wire', 'hookup', 'testing', 'finishing', 'ship_loose', 'assembly', 'sheet_metal')
		
		try:
			applicable_fields = self.applicable_labor_hours_per_material[material[0]]
		except:
			applicable_fields = labor_hour_fields
			
		for field in labor_hour_fields:
			if field not in applicable_fields:
				text_ctrl = ctrl(self, 'text:orders.labor_hours.{}'.format(field))
				text_ctrl.SetBackgroundColour(wx.Colour(125, 135, 140))
				text_ctrl.SetForegroundColour(wx.Colour(255, 250, 250))


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
			if field == None or field == 0:
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


	def on_click_toggle_status(self, event):
		record = db.query('''
			SELECT
				status
			FROM
				orders.root
			WHERE
				orders.root.id={}
			'''.format(self.id))
		
		current_status = record[0]
		
		if current_status == 'Valid':
			new_status = 'Canceled'
		else:
			new_status = 'Valid'
			
		db.update_order('orders.root', self.id, 'status', new_status)
			
		self.Freeze()
		self.reset_all()
		self.populate_all()
		self.Thaw()


	def on_click_specify_quote_number(self, event):
		record = db.query('''
			SELECT
				filemaker_quote
			FROM
				orders.root
			WHERE
				orders.root.id={}
			'''.format(self.id))
		
		current_quote_number = record[0]
		
		if not current_quote_number:
			current_quote_number = ''
		
		dialog =  wx.TextEntryDialog(self, 'Enter the Quote Number for this order:', 'Quote Number Input')
		dialog.SetValue(current_quote_number)
		
		if dialog.ShowModal() == wx.ID_OK:
			new_quote_number = dialog.GetValue()
			
			table = 'orders.root'
			field = 'filemaker_quote'

			records = db.query('''
				SELECT
					id
				FROM
					orders.root
				WHERE
					sales_order=(SELECT TOP 1 sales_order FROM orders.root WHERE id={})
				'''.format(self.id))
			
			for record in records:
				table_id = record
				db.update_order(table, table_id, field, new_quote_number)
			
			self.Freeze()
			self.reset_all()
			self.populate_all()
			self.Thaw()
			
		dialog.Destroy()


	def init_tabulated_data_tab(self):
		table_panel = ctrl(self, 'panel:tabulated_data')
		
		self.tabulated_data = TweakedGrid.TweakedGrid(table_panel)
		
		columns = list(db.get_table_column_names('orders.view_systems', presentable=False))
		
		for column_index, column in enumerate(columns):
			if '_spacer_' in column:
				column = ''
			columns[column_index] = '{}'.format(column)

		self.tabulated_data.CreateGrid(len(columns), 2)
		self.tabulated_data.SetRowLabelSize(0)
		self.tabulated_data.SetColLabelValue(0, 'Field')
		self.tabulated_data.SetColLabelValue(1, 'Value')
		
		for row_index, row_value in enumerate(columns):
			self.tabulated_data.SetCellValue(row_index, 0, row_value)
			self.tabulated_data.SetReadOnly(row_index, 0)
			self.tabulated_data.SetReadOnly(row_index, 1)
			self.tabulated_data.SetCellAlignment(row_index, 0, wx.ALIGN_RIGHT, wx.ALIGN_CENTRE)
		
		self.tabulated_data.AutoSize()
		self.tabulated_data.SetColSize(1, 400)
		self.tabulated_data.EnableDragRowSize(False)
		
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(self.tabulated_data, 1, wx.EXPAND)
		table_panel.SetSizer(sizer)
		
		table_panel.Layout()


	def reset_tabulated_data_tab(self):
		for row_index in range(self.tabulated_data.GetNumberRows()):
			self.tabulated_data.SetCellValue(row_index, 1, '')


	def populate_tabulated_data_tab(self):
		record = db.query('''
			SELECT
				*
			FROM
				orders.view_systems
			WHERE
				id={}
			'''.format(self.id))

		if record:
			#format all fields as strings
			formatted_record = []
			for field in record[0]:
				if field == None:
					field = ''
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%Y')
				else:
					field = '{}'.format(field)
					
				formatted_record.append(field)
		
			for row_index, row_value in enumerate(formatted_record):
				self.tabulated_data.SetCellValue(row_index, 1, row_value)


	def init_details_panel(self):
		ctrl(self, 'text:orders.misc.comments').SetBackgroundColour(ctrl(self, 'panel:main').GetBackgroundColour())
		
		ctrl(self, 'label:quote_label').Bind(wx.EVT_LEFT_DOWN, self.on_click_specify_quote_number)
		ctrl(self, 'label:quote').Bind(wx.EVT_LEFT_DOWN, self.on_click_specify_quote_number)
		
		ctrl(self, 'label:status_label').Bind(wx.EVT_LEFT_DOWN, self.on_click_toggle_status)
		ctrl(self, 'label:status').Bind(wx.EVT_LEFT_DOWN, self.on_click_toggle_status)


	def reset_details_panel(self):
		#return status colors to default
		ctrl(self, 'label:status').SetBackgroundColour(wx.NullColour)
		ctrl(self, 'label:status').SetForegroundColour(wx.NullColour)


	def populate_details_panel(self):
		record = db.query('''
			SELECT
				status,
				filemaker_quote,
				sales_order,
				item,
				production_order,
				material,
				hierarchy,
				model,
				description,
				serial,
				
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
				bpcs_family,

				comments
			FROM
				orders.root
			LEFT JOIN
				orders.misc ON orders.root.id = orders.misc.id
			WHERE
				orders.root.id={}
			'''.format(self.id))
		
		if not record:
			#if no root record was found for this Id, likely it was a bad idea got from searching a table like orders.changes
			# and then switching the table to orders.root where the Id locations are different in the search results list...
			# let's go ahead and bail so we don't accidently create a new order with this trash Id
			wx.MessageBox('An entry was not found with id={} in table orders.root.\n\nThis is likely caused by activating a search result in a table after switching to a new table to search in where the order id is not located in the same column as the other table. Redoing the search should solve the issue.'.format(self.id), 'Order Not Found', wx.OK | wx.ICON_INFORMATION)
			self.Close()
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

		status, filemaker_quote, sales_order, item, production_order, material, hierarchy, model, description, serial, \
		sold_to_name, sold_to_number, ship_to_name, ship_to_number, country, state, city, zip_code, address, \
		bpcs_sales_order, bpcs_line_up, bpcs_item, bpcs_family, comments = formatted_record
		
		#make a canceled order stand out
		if status == 'Canceled':
			ctrl(self, 'label:status').SetBackgroundColour((255, 0, 0))
			ctrl(self, 'label:status').SetForegroundColour((255, 255, 255))

		ctrl(self, 'label:status').SetLabel(status)
		ctrl(self, 'label:quote').SetLabel(filemaker_quote)
		ctrl(self, 'label:sales_order').SetLabel(sales_order)
		ctrl(self, 'label:item').SetLabel(item)
		ctrl(self, 'label:production_order').SetLabel(production_order)
		ctrl(self, 'label:material').SetLabel(material)
		ctrl(self, 'label:hierarchy').SetLabel(hierarchy)
		ctrl(self, 'label:model').SetLabel(model)
		ctrl(self, 'label:description').SetLabel(description)
		ctrl(self, 'label:serial').SetLabel(serial)

		ctrl(self, 'label:sold_to').SetLabel('{} ({})'.format(sold_to_name, sold_to_number))
		ctrl(self, 'label:ship_to').SetLabel('{} ({})'.format(ship_to_name, ship_to_number))
		ctrl(self, 'label:address').SetLabel(address)
		ctrl(self, 'label:city_state').SetLabel('{}, {} ({}) {}'.format(city, state, country, zip_code))
		
		#format the BPCS Sales Order and Line Up together
		bpcs_sales_order_and_line_up = '{}-{}'.format(bpcs_sales_order, bpcs_line_up)
		if bpcs_sales_order_and_line_up == '...-...':
			bpcs_sales_order_and_line_up = '...'

		if comments == '...':
			comments = ''
		ctrl(self, 'text:orders.misc.comments').SetValue(comments)

		ctrl(self, 'label:bpcs_item').SetLabel(bpcs_item)
		ctrl(self, 'label:bpcs_sales_order').SetLabel(bpcs_sales_order_and_line_up)
		ctrl(self, 'label:bpcs_family').SetLabel(bpcs_family)


	def on_click_auto_sign_up(self, event):
		#this will populate the user's name if none selected yet without opening the drop down
		choice_ctrl = event.GetEventObject()
		
		if choice_ctrl.GetStringSelection() == '':
			#choice_ctrl.SetStringSelection(gn.user)
			wx.CallAfter(choice_ctrl.SetStringSelection, gn.user)
		
		else:
			event.Skip()
			
		choice_ctrl.SetFocus()


	def init_responsibilities_tab(self):
		possible_signees = db.query('''
			SELECT
				name
			FROM
				dbo.employees
			WHERE
				department LIKE '%engineer%'
			ORDER BY
				name ASC
			''')
		
		possible_signees.insert(0, gn.user)
		possible_signees.insert(0, '')

		ctrl(self, 'choice:orders.responsibilities.applications_engineer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.applications_engineer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.design_engineer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.design_engineer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.mechanical_engineer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.mechanical_engineer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.mechanical_cad_designer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.mechanical_cad_designer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.electrical_engineer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.electrical_engineer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.electrical_cad_designer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.electrical_cad_designer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.structural_engineer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.structural_engineer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)

		ctrl(self, 'choice:orders.responsibilities.structural_cad_designer').AppendItems(possible_signees)
		ctrl(self, 'choice:orders.responsibilities.structural_cad_designer').Bind(wx.EVT_LEFT_DOWN, self.on_click_auto_sign_up)


		statuses = ['', 'Previewed', 'In Process', 'Reviewing', 'Complete', 'Hold']
		ctrl(self, 'choice:orders.responsibilities.applications_status').AppendItems(statuses)
		ctrl(self, 'choice:orders.responsibilities.design_status').AppendItems(statuses)
		ctrl(self, 'choice:orders.responsibilities.mechanical_status').AppendItems(statuses)
		ctrl(self, 'choice:orders.responsibilities.electrical_status').AppendItems(statuses)
		ctrl(self, 'choice:orders.responsibilities.structural_status').AppendItems(statuses)


	def reset_responsibilities_tab(self):
		ctrl(self, 'choice:orders.responsibilities.applications_engineer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.design_engineer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.mechanical_engineer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.mechanical_cad_designer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.electrical_engineer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.electrical_cad_designer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.structural_engineer').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.structural_cad_designer').SetStringSelection('')

		ctrl(self, 'choice:orders.responsibilities.applications_status').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.design_status').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.mechanical_status').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.electrical_status').SetStringSelection('')
		ctrl(self, 'choice:orders.responsibilities.structural_status').SetStringSelection('')


	def populate_responsibilities_tab(self):
		record = db.query('''
			SELECT
				applications_engineer,
				applications_status,
				design_engineer,
				design_status,
				mechanical_engineer,
				mechanical_cad_designer,
				mechanical_status,
				electrical_engineer,
				electrical_cad_designer,
				electrical_status,
				structural_engineer,
				structural_cad_designer,
				structural_status
			FROM
				orders.responsibilities
			WHERE
				id={}
			'''.format(self.id))
		
		if record:
			#format all fields as strings
			formatted_record = []
			for field in record[0]:
				if field == None:
					field = ''
					
				else:
					field = str(field)
					
				formatted_record.append(field)

			applications_engineer, applications_status, \
			design_engineer, design_status, \
			mechanical_engineer, mechanical_cad_designer, mechanical_status, \
			electrical_engineer, electrical_cad_designer, electrical_status, \
			structural_engineer, structural_cad_designer, structural_status = formatted_record

			#if we have an odd value in the database we have to insert it into the choice control
			# before we can set that control to that value
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.applications_engineer')
			
			if applications_engineer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(applications_engineer, 1)

			choice_ctrl.SetStringSelection(applications_engineer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.applications_status')
			
			if applications_status not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(applications_status, 1)

			choice_ctrl.SetStringSelection(applications_status)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.design_engineer')
			
			if design_engineer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(design_engineer, 1)

			choice_ctrl.SetStringSelection(design_engineer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.design_status')
			
			if design_status not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(design_status, 1)

			choice_ctrl.SetStringSelection(design_status)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.mechanical_engineer')
			
			if mechanical_engineer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(mechanical_engineer, 1)

			choice_ctrl.SetStringSelection(mechanical_engineer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.mechanical_cad_designer')
			
			if mechanical_cad_designer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(mechanical_cad_designer, 1)

			choice_ctrl.SetStringSelection(mechanical_cad_designer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.mechanical_status')
			
			if mechanical_status not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(mechanical_status, 1)

			choice_ctrl.SetStringSelection(mechanical_status)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.electrical_engineer')
			
			if electrical_engineer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(electrical_engineer, 1)

			choice_ctrl.SetStringSelection(electrical_engineer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.electrical_cad_designer')
			
			if electrical_cad_designer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(electrical_cad_designer, 1)

			choice_ctrl.SetStringSelection(electrical_cad_designer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.electrical_status')
			
			if electrical_status not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(electrical_status, 1)

			choice_ctrl.SetStringSelection(electrical_status)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.structural_engineer')
			
			if structural_engineer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(structural_engineer, 1)

			choice_ctrl.SetStringSelection(structural_engineer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.structural_cad_designer')
			
			if structural_cad_designer not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(structural_cad_designer, 1)

			choice_ctrl.SetStringSelection(structural_cad_designer)

			#---
			choice_ctrl = ctrl(self, 'choice:orders.responsibilities.structural_status')
			
			if structural_status not in choice_ctrl.GetStrings():
				choice_ctrl.Insert(structural_status, 1)

			choice_ctrl.SetStringSelection(structural_status)





	def init_changes_tab(self):
		column_names = ['Id', 'Table', 'Field', 'Previous Value', 'New Value', 'Who Changed', 'When Changed']

		for index, column_name in enumerate(column_names):
			ctrl(self, 'list:changes').InsertColumn(index, column_name)

	def reset_changes_tab(self):
		list_ctrl = ctrl(self, 'list:changes')
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

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


	def init_time_logs_tab(self):
		column_names = ['Id', 'Hours', 'Who Logged', 'When Logged', 'Tags']

		for index, column_name in enumerate(column_names):
			ctrl(self, 'list:time_logs').InsertColumn(index, column_name)

	def reset_time_logs_tab(self):
		list_ctrl = ctrl(self, 'list:time_logs')
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

	def populate_time_logs_tab(self):
		list_ctrl = ctrl(self, 'list:time_logs')
		list_ctrl.Freeze()
		
		records = db.query('''
			SELECT
				id,
				hours,
				who_logged,
				when_logged,
				tags
			FROM
				orders.time_logs
			WHERE
				order_id={}
			ORDER BY
				id DESC
			'''.format(self.id))
		
		#insert records into list
		for index, record in enumerate(records):
			id, hours, who_logged, when_logged, tags = record
			
			if tags == None:
				tags = ''

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(hours))
			list_ctrl.SetStringItem(index, 2, '{}'.format(who_logged))
			list_ctrl.SetStringItem(index, 3, '{}'.format(when_logged.strftime('%m/%d/%y %I:%M %p')))
			list_ctrl.SetStringItem(index, 4, '{}'.format(tags))

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)
			
			#cap column width at max 400
			if list_ctrl.GetColumnWidth(index) > 400:
				list_ctrl.SetColumnWidth(index, 400)
		
		#hide id column
		list_ctrl.SetColumnWidth(0, 0)
		
		list_ctrl.Thaw()




	def on_click_log_time(self, event):
		button = event.GetEventObject()
		button_name = button.Name.split(':')[1].replace('log_time_', '')

		LogTimeDialog(self, button_name)


	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()



class LogTimeDialog(wx.Dialog):
	def __init__(self, parent, button_name=None):
		#load frame XRC description
		pre = wx.PreDialog()
		res = xrc.XmlResource.Get() 
		res.LoadOnDialog(pre, parent, "dialog:log_time") 
		self.PostCreate(pre)
		self.SetIcon(wx.Icon(gn.resource_path('OrderManager.ico'), wx.BITMAP_TYPE_ICO))
		
		self.parent = parent
		self.button_name = button_name
		
		print 'self.button_name:', self.button_name
		
		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_dialog)
		self.Bind(wx.EVT_BUTTON, self.on_click_log_time, id=xrc.XRCID('button:log'))
		#self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.on_double_click_open_modify_time_log_dialog, id=xrc.XRCID('list:previous'))


		#set some toggle button defaults based on which "clock" was pressed
		if button_name == 'applications_engineer':
			ctrl(self, 'toggle:quoting').SetValue(True)

		elif button_name == 'design_engineer':
			#ctrl(self, 'toggle:planning').SetValue(True)
			ctrl(self, 'toggle:review').SetValue(True)
			ctrl(self, 'toggle:mechanical').SetValue(True)
			ctrl(self, 'toggle:electrical').SetValue(True)
			ctrl(self, 'toggle:structural').SetValue(True)

		elif button_name == 'mechanical_engineer':
			ctrl(self, 'toggle:review').SetValue(True)
			ctrl(self, 'toggle:mechanical').SetValue(True)

		elif button_name == 'mechanical_cad_designer':
			ctrl(self, 'toggle:cad_design').SetValue(True)
			ctrl(self, 'toggle:mechanical').SetValue(True)

		elif button_name == 'electrical_engineer':
			ctrl(self, 'toggle:review').SetValue(True)
			ctrl(self, 'toggle:electrical').SetValue(True)

		elif button_name == 'electrical_cad_designer':
			ctrl(self, 'toggle:cad_design').SetValue(True)
			ctrl(self, 'toggle:electrical').SetValue(True)

		elif button_name == 'structural_engineer':
			ctrl(self, 'toggle:review').SetValue(True)
			ctrl(self, 'toggle:structural').SetValue(True)

		elif button_name == 'structural_cad_designer':
			ctrl(self, 'toggle:cad_design').SetValue(True)
			ctrl(self, 'toggle:structural').SetValue(True)

		#populate previous user's time logs for this item
		list_ctrl = ctrl(self, 'list:previous')
		
		#create list headings
		column_names = ['Id', 'Hours', 'When Logged', 'Tags']
		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)
		
		records = db.query('''
			SELECT
				id,
				hours,
				when_logged,
				tags
			FROM
				orders.time_logs
			WHERE
				order_id={} AND 
				who_logged='{}'
			ORDER BY
				when_logged DESC
			'''.format(parent.id, gn.user.replace("'", "''")))
		
		total_hours = 0
		
		for log_index, log in enumerate(records):
			id, hours, when_logged, tags = log
			
			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(log_index, 0, '{}'.format(id))
			#list_ctrl.SetStringItem(log_index, 1, '{:.1f}'.format(hours))
			list_ctrl.SetStringItem(log_index, 1, '{}'.format(hours))
			list_ctrl.SetStringItem(log_index, 2, '{}'.format(when_logged.strftime('%m/%d/%y %I:%M %p')))
			list_ctrl.SetStringItem(log_index, 3, '{}'.format(tags))
			
			total_hours += hours

		for index, column_name in enumerate(column_names):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)

		#hide id column
		list_ctrl.SetColumnWidth(0, 0)

		ctrl(self, 'label:total_hours').SetLabel('{}'.format(total_hours))

		#misc
		self.ShowModal()


	def on_click_log_time(self, event):
		hours = ctrl(self, 'text:hours').GetValue()
		
		if hours == '':
			wx.MessageBox('Enter hours worked before logging time.', 'Hint', wx.OK | wx.ICON_WARNING)
			return
		
		tags = []
		if ctrl(self, 'toggle:quoting').GetValue(): tags.append('Quoting')
		if ctrl(self, 'toggle:planning').GetValue(): tags.append('Planning')
		if ctrl(self, 'toggle:cad_design').GetValue(): tags.append('CAD Design')
		if ctrl(self, 'toggle:review').GetValue(): tags.append('Review')
		if ctrl(self, 'toggle:mechanical').GetValue(): tags.append('Mechanical')
		if ctrl(self, 'toggle:electrical').GetValue(): tags.append('Electrical')
		if ctrl(self, 'toggle:structural').GetValue(): tags.append('Structural')
		if ctrl(self, 'toggle:revision').GetValue(): tags.append('Revision')
		if ctrl(self, 'toggle:shop_checkup').GetValue(): tags.append('Shop Checkup')
		
		#tags.sort()
		tags_string = ', '.join(tags)
		
		db.insert('orders.time_logs', (
			("order_id", self.parent.id),
			("who_logged", gn.user),
			("when_logged", 'CURRENT_TIMESTAMP'),
			("hours", hours),
			("tags", tags_string),)
		)

		self.parent.Freeze()
		self.parent.reset_all()
		self.parent.populate_all()
		self.parent.Thaw()

		self.Close()


	def on_close_dialog(self, event):
		print 'called on_close_dialog'
		self.Destroy()

