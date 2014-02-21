#!/usr/bin/env python
# -*- coding: utf8 -*-

import pyodbc
import sys
import os
import stat
import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import workdays
import datetime as dt
import General as gn
import Database as db
import Item


class DesignSchedulerFrame(wx.Frame):
	def __init__(self, parent):
		#load frame XRC description
		pre = wx.PreFrame()
		res = xrc.XmlResource.Get() 
		res.LoadOnFrame(pre, parent, "frame:design_scheduler") 
		self.PostCreate(pre)
		self.SetIcon(wx.Icon(gn.resource_path('OrderManager.ico'), wx.BITMAP_TYPE_ICO))
		
		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.Bind(wx.EVT_BUTTON, self.on_click_apply, id=xrc.XRCID('button:apply'))



		self.Bind(wx.EVT_LIST_ITEM_ACTIVATED , self.on_activated_order, id=xrc.XRCID('list:orders'))
		
		
		self.init_lists()
		self.populate_list()
		
		self.SetSize((1200, 700))
		self.Center()

		self.Show()


	def on_activated_order(self, event):
		selected_item = event.GetEventObject()
		table_id = selected_item.GetItem(selected_item.GetFirstSelected(), 0).GetText()

		if table_id != '':
			Item.ItemFrame(self, int(table_id))


	def init_lists(self):
		#a
		list_ctrl = ctrl(self, 'list:orders')
		
		column_names = ['Id', 'Sales Order', 'Item', 'Production Order', 'Material', 'Customer', '|',
						'When Got ProdOrd', 'Requested Release', '|', 'Date Locked', 'Old Planned Release', 'New Planned Release', 'Exception', '|', 'Old Suggested Start']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)


	def calculate_schedule(self, event=None):
		records = db.query('''
			SELECT
				id,
				material,
				date_created_on,
				date_requested_de_release
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				production_order IS NOT NULL AND
				date_requested_de_release IS NOT NULL AND
				date_planned_de_release_locked <> 1 AND
				status <> 'Canceled'
			ORDER BY
				date_requested_de_release, sales_order, item ASC
			''')
		
		self.schedule_data = {}
		
		for record in records:
			id, material, date_created_on, date_requested_de_release = record
			
			when_got_prodord = db.query('''
				SELECT TOP 1
					when_changed
				FROM
					orders.changes
				WHERE
					field = 'production_order' AND
					table_id = {}
				ORDER BY
					id DESC
				'''.format(id))
			
			if when_got_prodord:
				#since SAP export data is only updated every midnight, let's assume
				# that the actual change occured one working day before
				when_got_prodord = workdays.workday(when_got_prodord[0].date(), -1)
				
				if workdays.networkdays(when_got_prodord, date_requested_de_release.date()) >= 8:
					date_planned_de_release = date_requested_de_release
					
				else:
					#they did not give us 8 days between when it got a production order and when they want it released
					# so let's force that
					date_planned_de_release = workdays.workday(when_got_prodord, 8)
					
					#convert this date object to datetime
					date_planned_de_release = dt.datetime.combine(date_planned_de_release, dt.time())

			else:
				#we don't know when the order got the production order number, so give
				# the PC schedulers the benifit of the doubt
				if date_requested_de_release < date_created_on:
					date_planned_de_release = date_created_on
				
				else:
					date_planned_de_release = date_requested_de_release
					
				when_got_prodord = None

			
			#put code here to prevent planned date landing on a weekend
			self.schedule_data.update({id: (when_got_prodord, date_planned_de_release)})


	def on_click_apply(self, event=None):
		for id in self.schedule_data:
			when_got_prodord, new_date_planned_de_release = self.schedule_data[id]
			
			db.update_order('orders.target_dates', id, 'planned_de_release', new_date_planned_de_release, '{} (Auto)'.format(gn.user))

		self.Close()


	def populate_list(self, event=None):
		self.calculate_schedule()
		
		list_ctrl = ctrl(self, 'list:orders')
		list_ctrl.Freeze()
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

		records = db.query('''
			SELECT
				id,
				sales_order,
				item,
				production_order,
				material,
				sold_to_name,

				date_requested_de_release,
				date_planned_de_release,
				date_planned_de_release_locked,
				date_suggested_de_start
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				production_order IS NOT NULL AND
				date_requested_de_release IS NOT NULL AND
				status <> 'Canceled'
			ORDER BY
				date_planned_de_release, date_requested_de_release, sales_order, item ASC
			''')
		
	
		#insert records into list
		for index, record in enumerate(records):
			id, sales_order, item, production_order, material, sold_to_name, \
			date_requested_de_release, old_date_planned_de_release, date_planned_de_release_locked, date_suggested_de_start = record
			
			try:
				when_got_prodord, new_date_planned_de_release = self.schedule_data[id]

			except:
				when_got_prodord = '???'
				new_date_planned_de_release = '???'
			
			try:
				planned_exception = (new_date_planned_de_release - date_requested_de_release).days
			except:
				try:
					planned_exception = (old_date_planned_de_release - date_requested_de_release).days
				except:
					planned_exception = '???'
			

			if planned_exception < 0:
				planned_exception = 0



			#repack, format to string, then unpack
			record = (id, sales_order, item, production_order, material, sold_to_name, \
				when_got_prodord, date_requested_de_release, old_date_planned_de_release, new_date_planned_de_release, planned_exception, date_planned_de_release_locked, date_suggested_de_start)

			formatted_record = []
			for field in record:
				if field == None:
					field = ''
					
				elif isinstance(field, dt.datetime) or isinstance(field, dt.date):
					field = field.strftime('%m/%d/%Y')
					
				else:
					pass
					
				formatted_record.append(field)

			id, sales_order, item, production_order, material, sold_to_name, \
			when_got_prodord, date_requested_de_release, old_date_planned_de_release, new_date_planned_de_release, planned_exception, date_planned_de_release_locked, date_suggested_de_start = formatted_record
			


			if date_planned_de_release_locked == False:
				date_planned_de_release_locked = ''

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(sales_order))
			list_ctrl.SetStringItem(index, 2, '{}'.format(item))
			list_ctrl.SetStringItem(index, 3, '{}'.format(production_order))
			list_ctrl.SetStringItem(index, 4, '{}'.format(material))
			list_ctrl.SetStringItem(index, 5, '{}'.format(sold_to_name))
			
			list_ctrl.SetStringItem(index, 6, '{}'.format('|'))
			list_ctrl.SetStringItem(index, 7, '{}'.format(when_got_prodord))
			list_ctrl.SetStringItem(index, 8, '{}'.format(date_requested_de_release))
			
			list_ctrl.SetStringItem(index, 9, '{}'.format('|'))
			list_ctrl.SetStringItem(index, 10, '{}'.format(date_planned_de_release_locked))
			list_ctrl.SetStringItem(index, 11, '{}'.format(old_date_planned_de_release))
			list_ctrl.SetStringItem(index, 12, '{}'.format(new_date_planned_de_release))
			list_ctrl.SetStringItem(index, 13, '{}'.format(planned_exception))
			
			list_ctrl.SetStringItem(index, 14, '{}'.format('|'))
			list_ctrl.SetStringItem(index, 15, '{}'.format(date_suggested_de_start))


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



class SchedulingTab(object):
	def init_scheduling_tab(self):
		#Bindings
		###self.Bind(wx.EVT_CHOICE, self.on_choice_table, id=xrc.XRCID('choice:which_table'))
		
		
		self.Bind(wx.EVT_BUTTON, self.on_click_design_scheduler, id=xrc.XRCID('button:open_scheduler'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_proto_request_dates, id=xrc.XRCID('button:proto_request_dates'))
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_and_set_requested_de_release, id=xrc.XRCID('button:calc_and_set_requested_de_release'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_visualize_forecast, id=xrc.XRCID('button:visualize_forecast'))


	def get_engineering_capacity(self, date):
		if date.weekday() < 5:
			return 80
			
		else:
			return 0


	def on_click_design_scheduler(self, event):
		DesignSchedulerFrame(self)


	def on_click_visualize_forecast(self, event):
		ctrl(self, 'button:visualize_forecast').Disable()
		
		#make the excel read only
		os.chmod(gn.resource_path('VisualizeEtoForecast.xlsm'), stat.S_IREAD)

		today = dt.date.today()
		first_day_last_week = today - dt.timedelta(days=today.weekday()-2) - dt.timedelta(days=3) - dt.timedelta(weeks=1)
		last_day_of_report = first_day_last_week + dt.timedelta(weeks=16) #5)#dt.timedelta(days=7*5)

		with open(gn.resource_path("VisualizeEtoForecast.txt"), "w") as text_file:
			#write out headers
			headers = ['Date', 'Engineering Capacity', 'Requested In Process', 'Requested NOT In Process', 'Planned In Process', 'Planned NOT In Process', 'Start']
			text_file.write('{}\n'.format('`````'.join(headers)))
			
			#write out data
			for date_instance in gn.date_range(first_day_last_week, last_day_of_report):

				formatted_data = []

				capacity = self.get_engineering_capacity(date_instance)

				formatted_data.append(date_instance.strftime('%m/%d/%y'))
				formatted_data.append(str(capacity))
				
				#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				#requested
				records = db.query('''
					SELECT
						material
					FROM
						orders.view_systems
					WHERE
						date_actual_de_release IS NULL AND
						production_order IS NOT NULL AND
						status <> 'Canceled' AND
						
						date_requested_de_release = '{}'
					'''.format(date_instance))
				
				all_eto_hours = 0
				for material in records:
					all_eto_hours += sum(db.query("SELECT mean_hours FROM orders.material_eto_hour_estimates WHERE material='{}'".format(material)))

				records = db.query('''
					SELECT
						material
					FROM
						orders.view_systems
					WHERE
						date_actual_de_release IS NULL AND
						production_order IS NOT NULL AND
						status <> 'Canceled' AND
						
						date_requested_de_release = '{}' AND
						
						(design_status='In Process' OR 
						mechanical_status='In Process' OR 
						electrical_status='In Process' OR 
						structural_status='In Process')
					'''.format(date_instance))

				in_process_eto_hours = 0
				for material in records:
					in_process_eto_hours += sum(db.query("SELECT mean_hours FROM orders.material_eto_hour_estimates WHERE material='{}'".format(material)))

				formatted_data.append(str(in_process_eto_hours))
				formatted_data.append(str(all_eto_hours - in_process_eto_hours))

				#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				#planned
				records = db.query('''
					SELECT
						material
					FROM
						orders.view_systems
					WHERE
						date_actual_de_release IS NULL AND
						production_order IS NOT NULL AND
						status <> 'Canceled' AND
						
						date_planned_de_release = '{}'
					'''.format(date_instance))
				
				all_eto_hours = 0
				for material in records:
					all_eto_hours += sum(db.query("SELECT mean_hours FROM orders.material_eto_hour_estimates WHERE material='{}'".format(material)))

				records = db.query('''
					SELECT
						material
					FROM
						orders.view_systems
					WHERE
						date_actual_de_release IS NULL AND
						production_order IS NOT NULL AND
						status <> 'Canceled' AND
						
						date_planned_de_release = '{}' AND
						
						(design_status='In Process' OR 
						mechanical_status='In Process' OR 
						electrical_status='In Process' OR 
						structural_status='In Process')
					'''.format(date_instance))

				in_process_eto_hours = 0
				for material in records:
					in_process_eto_hours += sum(db.query("SELECT mean_hours FROM orders.material_eto_hour_estimates WHERE material='{}'".format(material)))

				formatted_data.append(str(in_process_eto_hours))
				formatted_data.append(str(all_eto_hours - in_process_eto_hours))


				#-------
				text_file.write('{}\n'.format('`````'.join(formatted_data)))

		os.startfile(gn.resource_path('VisualizeEtoForecast.xlsm'))

		ctrl(self, 'button:visualize_forecast').Enable()



	def on_click_calc_and_set_requested_de_release(self, event):
		records = db.query('''
			SELECT
				id,
				date_basic_start
			FROM
				orders.view_systems
			WHERE
				(date_actual_de_release IS NULL OR date_requested_de_release IS NULL) AND
				date_basic_start IS NOT NULL AND
				status <> 'Canceled'
			''')

		for record in records:
			id, date_basic_start = record
			
			calc_requested_de_release = workdays.workday(date_basic_start, -23)
			
			db.update_order('orders.target_dates', id, 'requested_de_release', calc_requested_de_release, '{} (Auto)'.format(gn.user))
			
		wx.MessageBox("{} requested_de_release fields have been calculated and set.".format(len(records)), 'Notice', wx.OK | wx.ICON_INFORMATION)


	def on_click_proto_request_dates(self, event):
		records = db.query('''
			SELECT
				id,
				
				CASE
					WHEN bpcs_sales_order IS NULL THEN sales_order
					WHEN sales_order IS NULL THEN bpcs_sales_order
					ELSE sales_order + '/' + bpcs_sales_order
				END AS sales_order,

				CASE
					WHEN bpcs_line_up IS NULL THEN item
					WHEN item IS NULL THEN bpcs_line_up
					ELSE item + '/' + bpcs_line_up
				END AS item,

				CASE
					WHEN bpcs_item IS NULL THEN production_order
					WHEN production_order IS NULL THEN bpcs_item
					ELSE production_order + '/' + bpcs_item
				END AS production_order,

				CASE
					WHEN bpcs_family IS NULL THEN material
					WHEN material IS NULL THEN bpcs_family
					ELSE material + '/' + bpcs_family
				END AS material,

				date_basic_start,
				date_requested_de_release
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				production_order IS NOT NULL AND
				material <> 'SPARTCOLS' AND
				status <> 'Canceled'
			ORDER BY
				date_basic_start ASC
			''')
		
		output = '	'.join(['Sales Order', 'Item', 'Production Order', 'When Got ProdOrd', 'Material', 'Basic Start', 'Ricky Request', 'Calculated Request'])
		output += '\n'
		
		for index, record in enumerate(records):
			id, sales_order, item, production_order, material, basic_start, ricky_request = record
			
			when_got_prod_ord = db.query('''
				SELECT TOP 1
					when_changed
				FROM
					orders.changes
				WHERE
					field = 'production_order' AND
					table_id = {}
				ORDER BY
					id DESC
				'''.format(id))


			
			try:
				calc_request = workdays.workday(basic_start, -23)
			except:
				calc_request = ''



			if when_got_prod_ord:
				when_got_prod_ord = when_got_prod_ord[0].strftime('%m/%d/%Y')
			else:
				when_got_prod_ord = ''

			if basic_start == None:
				basic_start = ''
				
			if ricky_request == None:
				ricky_request = ''
			
			try: basic_start = basic_start.strftime('%m/%d/%Y')
			except: pass

			try: ricky_request = ricky_request.strftime('%m/%d/%Y')
			except: pass

			try: calc_request = calc_request.strftime('%m/%d/%Y')
			except: pass

			output += '{}	'.format(sales_order)
			output += '{}	'.format(item)
			output += '{}	'.format(production_order)
			output += '{}	'.format(when_got_prod_ord)
			output += '{}	'.format(material)
			output += '{}	'.format(basic_start)
			output += '{}	'.format(ricky_request)
			output += '{}\n'.format(calc_request)
			
			
		ctrl(self, 'text:proto_request_dates').SetValue(output)

