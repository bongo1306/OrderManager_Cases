#!/usr/bin/env python
# -*- coding: utf8 -*-
version = '0.2'

import sys
import os
import traceback
import subprocess
import threading
import json
import shutil
import workdays

import wx				#wxWidgets used as the GUI
from wx import xrc		#allows the loading and access of xrc file (xml) that describes GUI
ctrl = xrc.XRCCTRL		#define a shortined function name (just for convienience)
import wx.lib.agw.advancedsplash as AS
#import wx.lib.inspection
import wx.richtext as rt

import ConfigParser #for reading local config data (*.cfg)

import datetime as dt

import BetterListCtrl
import TextCtrlDbLinker
import CheckboxCtrlDbLinker
import ChoiceCtrlDbLinker
import ComboCtrlDbLinker

import General as gn
import Database as db
import Search
import Item


def check_for_updates():
	try:
		with open(os.path.join(gn.updates_dir, "releases.json")) as file:
			releases = json.load(file)
			
			latest_version = releases[0]['version']
			
			if version != latest_version:
				wx.CallAfter(open_software_update_frame)
	except Exception as e:
		print 'Failed update check:', e

def open_software_update_frame():
	SoftwareUpdateFrame(None)



class SoftwareUpdateFrame(wx.Frame):
	def __init__(self, parent):
		#load frame XRC description
		pre = wx.PreFrame()
		res = xrc.XmlResource.Get()  
		res.LoadOnFrame(pre, parent, "frame:software_update") 
		self.PostCreate(pre)
		self.SetIcon(wx.Icon(gn.resource_path('OrderManager.ico'), wx.BITMAP_TYPE_ICO))
		
		#read in update text data
		with open(os.path.join(gn.updates_dir, "releases.json")) as file:
			releases = json.load(file)
			
		latest_version = releases[0]['version']
		self.install_filename = releases[0]['installer filename']

		it_is_mandatory = False

		#build up what changed text
		changes_panel = ctrl(self, 'panel:changes')
		richtext_ctrl = rt.RichTextCtrl(changes_panel, style=wx.VSCROLL|wx.HSCROLL|wx.TE_READONLY)
		sizer = wx.BoxSizer(wx.VERTICAL)
		sizer.Add(richtext_ctrl, 1, wx.EXPAND)
		changes_panel.SetSizer(sizer)
		changes_panel.Layout()

		for release in releases:
			if float(version) < float(release['version']):

				richtext_ctrl.BeginBold()
				richtext_ctrl.WriteText('v{} - {}'.format(release['version'], release['release date']))
				richtext_ctrl.EndBold()
				richtext_ctrl.Newline()
				
				richtext_ctrl.BeginStandardBullet('*', 50, 30)
				
				for change in release['changes']:
					richtext_ctrl.WriteText('{}'.format(change))
					richtext_ctrl.Newline()
				
				richtext_ctrl.EndStandardBullet()
				
				if release['mandatory'] == True:
					it_is_mandatory = True


		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.Bind(wx.EVT_BUTTON, self.on_click_cancel, id=xrc.XRCID('button:not_now'))
		self.Bind(wx.EVT_BUTTON, self.on_click_update, id=xrc.XRCID('button:update'))
		
		#misc
		ctrl(self, 'label:intro_version').SetLabel("A new version of the OrderManager software was found on the server: v{}".format(latest_version))
		ctrl(self, 'button:update').SetFocus()
		
		if it_is_mandatory == False:
			ctrl(self, 'button:not_now').Enable()
			ctrl(self, 'label:mandatory').Hide()
		
		self.Show()



	def on_click_cancel(self, event):
		self.Close()


	def on_click_update(self, event):
		print 'copy install file over, open it, and close this program'
		#create a dialog to show log of what's goin on
		dialog = wx.Dialog(self, id = wx.ID_ANY, title = u"Starting Update...", pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER)
		dialog.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)
		dialog.SetFont(wx.Font(10, 70, 90, 90, False, wx.EmptyString))
		bSizer53 = wx.BoxSizer(wx.VERTICAL)
		dialog.m_panel22 = wx.Panel(dialog, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
		bSizer54 = wx.BoxSizer(wx.VERTICAL)
		bSizer55 = wx.BoxSizer(wx.VERTICAL)
		dialog.text_notice = wx.TextCtrl(dialog.m_panel22, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size(350,120), wx.TE_DONTWRAP|wx.TE_MULTILINE|wx.TE_READONLY)
		bSizer55.Add(dialog.text_notice, 1, wx.ALL|wx.EXPAND, 5)
		bSizer54.Add(bSizer55, 1, wx.EXPAND, 5)
		dialog.m_panel22.SetSizer(bSizer54)
		dialog.m_panel22.Layout()
		bSizer54.Fit(dialog.m_panel22)
		bSizer53.Add(dialog.m_panel22, 1, wx.EXPAND |wx.ALL, 5)
		dialog.SetSizer(bSizer53)
		dialog.Layout()
		bSizer53.Fit(dialog)
		dialog.Centre(wx.BOTH)
		dialog.Show()
		
		dialog.text_notice.AppendText('Downloading install file from server... \n')
		dialog.text_notice.AppendText('This may take several minutes depending\n')
		dialog.text_notice.AppendText('on your connection... ')
		wx.Yield()
		
		try:
			source_filepath = os.path.join(gn.updates_dir, self.install_filename)
			destination_filepath = gn.resource_path(self.install_filename)
			
			shutil.copy2(source_filepath, destination_filepath)
			
		except Exception as e:
			dialog.text_notice.AppendText('[FAIL]\n')
			dialog.text_notice.AppendText('ERROR: {}\n'.format(e))
			dialog.text_notice.AppendText('\nSoftware update failed.\n')
			wx.Yield()
			return
		

		dialog.text_notice.AppendText('[OK]\n')
		dialog.text_notice.AppendText('Opening install file... ')
		wx.Yield()
		
		try:
			os.startfile(destination_filepath)
			
		except Exception as e:
			dialog.text_notice.AppendText('[FAIL]\n')
			dialog.text_notice.AppendText('ERROR: {}\n'.format(e))
			dialog.text_notice.AppendText('\nSoftware update failed.\n')
			wx.Yield()
			return
			
		dialog.text_notice.AppendText('[OK]\n')
		dialog.text_notice.AppendText('Shutting down this program... ')
		wx.Yield()
		
		try:
			try: gn.splash_frame.Close()
			except Exception as e: print 'splash', e

			try: gn.login_frame.Close()
			except Exception as e: print 'login', e

			try: gn.main_frame.Close()
			except Exception as e: print 'main', e
			
			self.Close()
			
		except Exception as e:
			dialog.text_notice.AppendText('[FAIL]\n')
			dialog.text_notice.AppendText('ERROR: {}\n'.format(e))
			dialog.text_notice.AppendText('\nSoftware update failed.\n')
			wx.Yield()
			return
		

	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()



class LoginFrame(wx.Frame):
	def __init__(self, parent):
		#load frame XRC description
		pre = wx.PreFrame()
		res = xrc.XmlResource.Get() 
		res.LoadOnFrame(pre, parent, "frame:login") 
		self.PostCreate(pre)
		self.SetIcon(wx.Icon(gn.resource_path('OrderManager.ico'), wx.BITMAP_TYPE_ICO))
		
		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.Bind(wx.EVT_BUTTON, self.on_click_login, id=xrc.XRCID('button:log_in'))
		#self.Bind(wx.EVT_BUTTON, self.on_click_create_user, id=xrc.XRCID('button:create_user'))

		#misc
		self.SetTitle('OrderManager v{} Login'.format(version))

		#populate user selection box
		names = db.query("SELECT name FROM employees WHERE activated = 1 ORDER BY name ASC")
		ctrl(self, 'choice:name').AppendItems(names)
		
		#create OrderManager config file if it doesn't exits yet
		if not os.path.exists(gn.resource_path('OrderManager.cfg')):
			with open(gn.resource_path('OrderManager.cfg'), "w") as text_file:
				text_file.write('[Application]\nlogin_name = \npassword = \nremember_password = True\nprevious_main_tab = \nprevious_item_tab = \nprevious_search_table =\nsearch_criteria_alphabetical = False')
				
		#default the login name to the last entered name
		login_name = ''
		remember_password = ''
		password = ''
		config = ConfigParser.ConfigParser()
		if config.read(gn.resource_path('OrderManager.cfg')):
			login_name = config.get('Application', 'login_name')
			remember_password = config.get('Application', 'remember_password')
			password = config.get('Application', 'password')

		if login_name != '':
			ctrl(self, 'choice:name').SetStringSelection(login_name)

		if remember_password == 'True':
			ctrl(self, 'check:remember').SetValue(True)
			ctrl(self, 'text:password').SetValue(password)

		#put focus on password box
		ctrl(self, 'text:password').SetFocus()
		self.Show()
		
		#check for updates!
		threading.Thread(target=check_for_updates).start()
		
		#gn.splash_frame.Close()
		
		##############
		#self.on_click_login()


	def on_click_login(self, event=None):
		#get entered user and password values from form fields
		selected_user = ctrl(self, 'choice:name').GetStringSelection()
		entered_password = ctrl(self, 'text:password').GetValue()
		remember_checked = ctrl(self, 'check:remember').GetValue()

		#field validation
		if selected_user == '':
			wx.MessageBox('You must select a user.', 'Login failed')
			return

		#get password from selected user name
		real_password = db.query("SELECT TOP 1 password FROM employees WHERE name = '{}'".format(selected_user.replace("'", "''")))[0]

		#if real_password == '':
		#	SetPasswordDialog(None)
		#	return

		#check credentials
		if entered_password.upper() == real_password.upper():
			#set the default login name in config file to the one just entered
			config = ConfigParser.ConfigParser()
			config.read(gn.resource_path('OrderManager.cfg'))
			config.set('Application', 'login_name', selected_user)
			
			if remember_checked == True:
				config.set('Application', 'remember_password', 'True')
				config.set('Application', 'password', entered_password)
			else:
				config.set('Application', 'remember_password', 'False')
				config.set('Application', 'password', '')
				
			with open(gn.resource_path('OrderManager.cfg'), 'w+') as configfile:
				config.write(configfile)

			gn.user = selected_user
			gn.main_frame = MainFrame(None)

			self.Close()
			
		else:
			wx.MessageBox('Invalid password for user %s.' % selected_user, 'Login failed')


	#def on_click_create_user(self, event):
	#	CreateNewUserFrame(self)


	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()



class MainFrame(wx.Frame, Search.SearchTab):
	def __init__(self, parent):
		#super(MainFrame, self).__init__(self)
		
		#load frame XRC description
		pre = wx.PreFrame()
		res = xrc.XmlResource.Get() 
		res.LoadOnFrame(pre, parent, "frame:main") 
		self.PostCreate(pre)
		self.SetIcon(wx.Icon(gn.resource_path('OrderManager.ico'), wx.BITMAP_TYPE_ICO))
		
		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		#self.Bind(wx.EVT_BUTTON, self.on_click_login, id=xrc.XRCID('button:log_in'))
		#self.Bind(wx.EVT_BUTTON, self.on_click_create_user, id=xrc.XRCID('button:create_user'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_proto_request_dates, id=xrc.XRCID('button:proto_request_dates'))
		
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_and_set_requested_de_release, id=xrc.XRCID('button:calc_and_set_requested_de_release'))
		


		#misc
		self.SetTitle('OrderManager v{} - Logged in as {}'.format(version, gn.user))
		#OrderScheduler v{} - Logged in as{} {}'.format(version, self.user.split(',')[-1], self.user.split(',')[0]))

		self.init_search_tab()
		self.init_lists()
		self.init_reports()
		
		self.refresh_list_lacking_quote_ae()
		self.refresh_list_unreleased_ae()
		self.refresh_list_unreleased_de()
		self.refresh_list_exceptions_de()
		self.refresh_list_sent_to_mmg()


		self.Show()
		
		#Item.ItemFrame(self, id=12588966)
		#Item.ItemFrame(self, id=12585270)
		#Item.ItemFrame(self, id=12585556)
		#Item.ItemFrame(self, id=12588283)
		#Item.ItemFrame(self, id=12587666)
		#Item.ItemFrame(self, id=12585385)


	def on_click_calc_and_set_requested_de_release(self, event):
		records = db.query('''
			SELECT
				id,
				date_basic_start
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				date_basic_start IS NOT NULL AND
				material <> 'SPARTCOLS' AND
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




	def on_activated_order(self, event):
		selected_item = event.GetEventObject()
		table_id = selected_item.GetItem(selected_item.GetFirstSelected(), 0).GetText()

		if table_id != '':
			Item.ItemFrame(self, int(table_id))


	def init_reports(self):
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_de_release_stats, id=xrc.XRCID('button:calc_de_release_stats'))
		
		today = dt.date.today()
		first_day_last_week = today - dt.timedelta(days=today.weekday()-2) - dt.timedelta(days=3) - dt.timedelta(weeks=1)
		last_day_last_week = first_day_last_week + dt.timedelta(days=6)
		
		ctrl(self, 'date:de_release_stats_from').SetValue(wx.DateTimeFromDMY(first_day_last_week.day, first_day_last_week.month-1, first_day_last_week.year))		
		ctrl(self, 'date:de_release_stats_to').SetValue(wx.DateTimeFromDMY(last_day_last_week.day, last_day_last_week.month-1, last_day_last_week.year))		


	def on_click_calc_de_release_stats(self, event):
		start_date = gn.wxdate2pydate(ctrl(self, 'date:de_release_stats_from').GetValue())
		end_date = gn.wxdate2pydate(ctrl(self, 'date:de_release_stats_to').GetValue())

		report = ctrl(self, 'text:de_release_stats')
		report.SetValue('')
		
		#report.AppendText("Considering items DE released from {} to {},\n\n".format(start_date.strftime('%m/%d/%y'), end_date.strftime('%m/%d/%y')))

		rel_items_data = db.query('''
			SELECT
				id,
				production_order,
				date_requested_de_release,
				date_planned_de_release,
				date_actual_de_release,
				material,
				hours_standard
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release >= '{}' AND
				date_actual_de_release <= '{} 23:59:59'
			ORDER BY
				date_actual_de_release
			'''.format(start_date, end_date))

		est_de_hours = 0
		act_de_hours = 0

		act_std_hours = 0
		est_std_hours = 0

		ontime_by_request = 0
		ontime_by_planned = 0
		
		released_ahead_of_schedule = 0
		released_ahead_of_schedule_std_hours = 0
		
		items_with_no_logged_time = []
		
		items_rel_late = []
		
		for rel_item_data in rel_items_data:
			id, production_order, date_requested_de_release, date_planned_de_release, date_actual_de_release, material, hours_standard = rel_item_data

			try:
				if date_actual_de_release <= date_requested_de_release:
					ontime_by_request += 1
			except Exception as e:
				print e
				
			try:
				if date_actual_de_release <= date_planned_de_release:
					ontime_by_planned += 1
				else:
					items_rel_late.append("{} - released {} working days late".format(production_order, workdays.networkdays(date_planned_de_release, date_actual_de_release)-1))
					
			except Exception as e:
				#items_rel_late.append("{} - no Com Rel Date".format(item))
				#thinking positively, just assume on time if item was never scheduled to begin with.
				ontime_by_planned += 1
				print e
			
			est_de_hours += sum(db.query("SELECT mean_hours FROM family_hour_estimates WHERE family='{}'".format(material)))
			
			this_act_de_hours = sum(db.query("SELECT hours FROM orders.time_logs WHERE order_id={}".format(id)))
			if this_act_de_hours == 0:
				#print item
				items_with_no_logged_time.append((id, production_order))
			else:
				act_de_hours += this_act_de_hours
				
			this_est_std_hours = 0
			if hours_standard == 0:
				try:
					this_est_std_hours = db.query("SELECT mean_hours FROM std_family_hour_estimates WHERE family='{}'".format(material))[0]
				except:
					print 'failed to gather this_est_std_hours for material:', material
					this_est_std_hours = 0
				est_std_hours += this_est_std_hours
			else:
				act_std_hours += hours_standard
			
			try:
				#if date_planned_de_release > dt.datetime.combine(end_date, dt.time()):
				if date_planned_de_release > date_actual_de_release + dt.timedelta(days=6-date_actual_de_release.weekday()):
					#print item, '{}	{}'.format(date_planned_de_release, (hours_standard + this_est_std_hours))
					released_ahead_of_schedule += 1
					released_ahead_of_schedule_std_hours += (hours_standard + this_est_std_hours) #one of them is zero so it's koo
			except Exception as e:
				print 'checking if date_planned_de_release > end_date for {}'.format(production_order), e
				

		report.AppendText("{} items were released by DE between {} and {},\n\n".format(len(rel_items_data), start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y')))
		
		if rel_items_data:
			report.AppendText("     {} items or {:.0f}% were released on time based on date_requested_de_release\n".format(ontime_by_request, ontime_by_request/float(len(rel_items_data))*100))
			report.AppendText("     {} items or {:.0f}% were released on time based on date_planned_de_release\n\n".format(ontime_by_planned, ontime_by_planned/float(len(rel_items_data))*100))

			report.AppendText("     {:.0f} hours estimated worked by DE based on our family hour estimates\n".format(est_de_hours))
			report.AppendText("     {:.0f} hours actually logged by DE\n\n".format(act_de_hours))

			report.AppendText("     {:.0f} standard hours released, {:.0f} of which are actually known while {:.0f} are estimated\n".format(est_std_hours+float(act_std_hours), act_std_hours, est_std_hours))
			report.AppendText("     {} items where released ahead of schedule, accounting for {:.0f} of the {:.0f} standard hours\n\n\n".format(released_ahead_of_schedule, released_ahead_of_schedule_std_hours, est_std_hours+float(act_std_hours)))

			if items_with_no_logged_time:
				report.AppendText("The following items have no time logged for them:\n")
				
				for item_with_no_logged_time in items_with_no_logged_time:
					
					ord_id, prod_ord = item_with_no_logged_time
					
					#print item_with_no_logged_time
					project_lead = db.query("SELECT design_engineer FROM orders.responsibilities WHERE id={}".format(ord_id))
					
					if project_lead:
						report.AppendText("     {:<9} - Project Lead is {}\n".format(prod_ord, project_lead[0]))
					else:
						report.AppendText("     {:<9} - No Project Lead... Maybe AE did it?\n".format(prod_ord))
						
			if items_rel_late:
				report.AppendText("\nThe following items are late according to date_planned_de_release:\n")
				
				for entry in items_rel_late:
					report.AppendText("     {}\n".format(entry))
		
		report.AppendText("\n\nElvis, remember this is still using dbo.family_hour_estimates and dbo.std_family_hour_estimates...")




	def init_lists(self):
		#applications lacking quote number
		list_ctrl = ctrl(self, 'list:lacking_quote_ae')

		list_ctrl.printer_paper_type = wx.PAPER_11X17
		list_ctrl.printer_header = 'Orders Lacking Quote Numbers'
		list_ctrl.printer_font_size = 8
		
		self.Bind(wx.EVT_LIST_ITEM_ACTIVATED , self.on_activated_order, id=xrc.XRCID('list:lacking_quote_ae'))
		self.Bind(wx.EVT_BUTTON, self.refresh_list_lacking_quote_ae, id=xrc.XRCID('button:refresh_lacking_quote_ae'))
		self.Bind(wx.EVT_BUTTON, list_ctrl.filter_list, id=xrc.XRCID('button:filter_lacking_quote_ae')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.export_list, id=xrc.XRCID('button:export_lacking_quote_ae')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.print_list, id=xrc.XRCID('button:print_lacking_quote_ae')) 
		
		column_names = ['Id', 'Sales Order', 'Item', 'Production Order', 'Material', 'Customer', 'Std Hours', 
						'Requested Release', 'Planned Release',
						'Applications Lead', 'Applications Status']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)


		#applications unreleased
		list_ctrl = ctrl(self, 'list:unreleased_ae')

		list_ctrl.printer_paper_type = wx.PAPER_11X17
		list_ctrl.printer_header = 'AE Unreleased Schedule'
		list_ctrl.printer_font_size = 8
		
		self.Bind(wx.EVT_LIST_ITEM_ACTIVATED , self.on_activated_order, id=xrc.XRCID('list:unreleased_ae'))
		self.Bind(wx.EVT_BUTTON, self.refresh_list_unreleased_ae, id=xrc.XRCID('button:refresh_unreleased_ae'))
		self.Bind(wx.EVT_BUTTON, list_ctrl.filter_list, id=xrc.XRCID('button:filter_unreleased_ae')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.export_list, id=xrc.XRCID('button:export_unreleased_ae')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.print_list, id=xrc.XRCID('button:print_unreleased_ae')) 
		
		column_names = ['Id', 'Quote', 'Sales Order', 'Item', 'Production Order', 'Material', 'Customer', 'Std Hours', 
						'Requested Release', 'Planned Release',
						'Applications Lead', 'Applications Status']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)


		#design unreleased
		list_ctrl = ctrl(self, 'list:unreleased_de')

		list_ctrl.printer_paper_type = wx.PAPER_11X17
		list_ctrl.printer_header = 'DE Unreleased Schedule'
		list_ctrl.printer_font_size = 8
		
		self.Bind(wx.EVT_LIST_ITEM_ACTIVATED , self.on_activated_order, id=xrc.XRCID('list:unreleased_de'))
		self.Bind(wx.EVT_BUTTON, self.refresh_list_unreleased_de, id=xrc.XRCID('button:refresh_unreleased_de'))
		self.Bind(wx.EVT_BUTTON, list_ctrl.filter_list, id=xrc.XRCID('button:filter_unreleased_de')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.export_list, id=xrc.XRCID('button:export_unreleased_de')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.print_list, id=xrc.XRCID('button:print_unreleased_de')) 
		
		column_names = ['Id', 'Sales Order', 'Item', 'Production Order', 'Material', 'Customer', 'Std Hours', 
						'Requested Release', 'Planned Release', 'Suggested Start', 
						'Design Lead',
						'Design Status',
						'Mechanical Status',
						'Electrical Status',
						'Structural Status',
						'Mechanical Engineer',
						'Electrical Engineer',
						'Structural Engineer',
						'Mechanical CAD Designer', 
						'Electrical CAD Designer', 
						'Structural CAD Designer']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)

		#design release exceptions
		list_ctrl = ctrl(self, 'list:exceptions_de')

		list_ctrl.printer_paper_type = wx.PAPER_11X17
		list_ctrl.printer_header = 'DE Release Exceptions'
		list_ctrl.printer_font_size = 8
		
		self.Bind(wx.EVT_LIST_ITEM_ACTIVATED , self.on_activated_order, id=xrc.XRCID('list:exceptions_de'))
		self.Bind(wx.EVT_BUTTON, self.refresh_list_exceptions_de, id=xrc.XRCID('button:refresh_exceptions_de'))
		self.Bind(wx.EVT_BUTTON, list_ctrl.filter_list, id=xrc.XRCID('button:filter_exceptions_de')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.export_list, id=xrc.XRCID('button:export_exceptions_de')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.print_list, id=xrc.XRCID('button:print_exceptions_de')) 
		
		column_names = ['Id', 'Sales Order', 'Item', 'Production Order', 'Material', 'Customer', 'Std Hours', 
						'Requested Release', 'Planned Release', 'Days Off By', 
						'Design Lead',
						'Design Status',
						'Mechanical Status',
						'Electrical Status',
						'Structural Status',
						'Mechanical Engineer',
						'Electrical Engineer',
						'Structural Engineer',
						'Mechanical CAD Designer', 
						'Electrical CAD Designer', 
						'Structural CAD Designer']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)

		#recently sent to mmg
		list_ctrl = ctrl(self, 'list:sent_to_mmg')

		#list_ctrl.printer_paper_type = wx.PAPER_11X17
		list_ctrl.printer_header = 'Recently Sent to MMG'
		list_ctrl.printer_font_size = 8
		
		self.Bind(wx.EVT_LIST_ITEM_ACTIVATED , self.on_activated_order, id=xrc.XRCID('list:sent_to_mmg'))
		self.Bind(wx.EVT_BUTTON, self.refresh_list_sent_to_mmg, id=xrc.XRCID('button:refresh_sent_to_mmg'))
		self.Bind(wx.EVT_BUTTON, list_ctrl.filter_list, id=xrc.XRCID('button:filter_sent_to_mmg')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.export_list, id=xrc.XRCID('button:export_sent_to_mmg')) 
		self.Bind(wx.EVT_BUTTON, list_ctrl.print_list, id=xrc.XRCID('button:print_sent_to_mmg')) 
		
		column_names = ['Id', 'Sales Order', 'Item', 'Production Order', 'Material', 'Customer',
						'Filename', 'When Sent']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)


	def refresh_list_lacking_quote_ae(self, event=None):
		list_ctrl = ctrl(self, 'list:lacking_quote_ae')
		list_ctrl.Freeze()
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

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

				sold_to_name,
				hours_standard,

				date_requested_ae_release,
				date_planned_ae_release,

				applications_engineer,
				applications_status
			FROM
				orders.view_systems
			WHERE
				date_actual_ae_release IS NULL AND
				status <> 'Canceled' AND
				quote IS NULL
			ORDER BY
				sales_order, item ASC
			''')
		
		#insert records into list
		index = -1
		for record in records:
			#format all fields as strings
			formatted_record = []
			for field in record:
				if field == None:
					field = ''
					
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%Y')
					
				else:
					pass
					
				formatted_record.append(field)

			id, sales_order, item, production_order, material, sold_to_name, hours_standard, \
			date_requested_ae_release, date_planned_ae_release, \
			applications_engineer, applications_status = formatted_record

			#only display orders with CMATs that AE cares about
			if material not in ('CDA', 'CPP', 'CSS', 'CTL', 'CVS', 'DBV', 'DSP', 'DSS', 'DSSIIX', \
							'FAH', 'FAV', 'FAX', 'HPM', 'HVS', 'MISC', 'NH2', 'NHS', 'NV2', \
							'NX2', 'OHD', 'OHN', 'OHS', 'OHW', 'ONH', 'ONV', 'PSM', 'RCA', \
							'RHD', 'SHIP_LOOSE', 'WEE', 'WEH', 'WEM'):
				continue
			
			index += 1

			#remove the decimals from the std hours
			try:
				hours_standard = round(hours_standard, 1)
			except:
				pass

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(sales_order))
			list_ctrl.SetStringItem(index, 2, '{}'.format(item))
			list_ctrl.SetStringItem(index, 3, '{}'.format(production_order))
			list_ctrl.SetStringItem(index, 4, '{}'.format(material))
			list_ctrl.SetStringItem(index, 5, '{}'.format(sold_to_name))
			list_ctrl.SetStringItem(index, 6, '{}'.format(hours_standard))

			list_ctrl.SetStringItem(index, 7, '{}'.format(date_requested_ae_release))
			list_ctrl.SetStringItem(index, 8, '{}'.format(date_planned_ae_release))

			list_ctrl.SetStringItem(index, 9, '{}'.format(applications_engineer))
			list_ctrl.SetStringItem(index, 10, '{}'.format(applications_status))

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)
			
			#cap column width at max 400
			if list_ctrl.GetColumnWidth(index) > 400:
				list_ctrl.SetColumnWidth(index, 400)
		
		#hide id column
		list_ctrl.SetColumnWidth(0, 0)
		
		list_ctrl.Thaw()



	def refresh_list_unreleased_ae(self, event=None):
		list_ctrl = ctrl(self, 'list:unreleased_ae')
		list_ctrl.Freeze()
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

		records = db.query('''
			SELECT
				id,
				quote,

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

				sold_to_name,
				hours_standard,

				date_requested_ae_release,
				date_planned_ae_release,

				applications_engineer,
				applications_status
			FROM
				orders.view_systems
			WHERE
				date_actual_ae_release IS NULL AND
				status <> 'Canceled' AND
				quote IS NOT NULL
			ORDER BY
				date_planned_ae_release, date_requested_ae_release, sales_order, item ASC
			''')
		
		#insert records into list
		for index, record in enumerate(records):
			#format all fields as strings
			formatted_record = []
			for field in record:
				if field == None:
					field = ''
					
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%Y')
					
				else:
					pass
					
				formatted_record.append(field)

			id, quote, sales_order, item, production_order, material, sold_to_name, hours_standard, \
			date_requested_ae_release, date_planned_ae_release, \
			applications_engineer, applications_status = formatted_record
			
			#remove the decimals from the std hours
			try:
				hours_standard = round(hours_standard, 1)
			except:
				pass

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(quote))
			list_ctrl.SetStringItem(index, 2, '{}'.format(sales_order))
			list_ctrl.SetStringItem(index, 3, '{}'.format(item))
			list_ctrl.SetStringItem(index, 4, '{}'.format(production_order))
			list_ctrl.SetStringItem(index, 5, '{}'.format(material))
			list_ctrl.SetStringItem(index, 6, '{}'.format(sold_to_name))
			list_ctrl.SetStringItem(index, 7, '{}'.format(hours_standard))

			list_ctrl.SetStringItem(index, 8, '{}'.format(date_requested_ae_release))
			list_ctrl.SetStringItem(index, 9, '{}'.format(date_planned_ae_release))

			list_ctrl.SetStringItem(index, 10, '{}'.format(applications_engineer))
			list_ctrl.SetStringItem(index, 11, '{}'.format(applications_status))

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)
			
			#cap column width at max 400
			if list_ctrl.GetColumnWidth(index) > 400:
				list_ctrl.SetColumnWidth(index, 400)
		
		#hide id column
		list_ctrl.SetColumnWidth(0, 0)
		
		list_ctrl.Thaw()


	
	def refresh_list_unreleased_de(self, event=None):
		list_ctrl = ctrl(self, 'list:unreleased_de')
		list_ctrl.Freeze()
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

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

				sold_to_name,
				hours_standard,

				date_requested_de_release,
				date_planned_de_release,
				date_suggested_de_start,

				design_engineer,
				design_status,
				mechanical_status,
				electrical_status,
				structural_status,
				mechanical_engineer,
				electrical_engineer,
				structural_engineer,
				mechanical_cad_designer,
				electrical_cad_designer,
				structural_cad_designer
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				production_order IS NOT NULL AND
				status <> 'Canceled'
			ORDER BY
				date_requested_de_release, sales_order, item ASC
			''')
		
		#insert records into list
		for index, record in enumerate(records):
			#format all fields as strings
			formatted_record = []
			for field in record:
				if field == None:
					field = ''
					
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%Y')
					
				else:
					pass
					
				formatted_record.append(field)

			id, sales_order, item, production_order, material, sold_to_name, hours_standard, \
			date_requested_de_release, date_planned_de_release, date_suggested_de_start, \
			design_engineer, design_status, mechanical_status, electrical_status, structural_status, \
			mechanical_engineer, electrical_engineer, structural_engineer, \
			mechanical_cad_designer, electrical_cad_designer, structural_cad_designer = formatted_record
			
			#remove the decimals from the std hours
			try:
				hours_standard = round(hours_standard, 1)
			except:
				pass

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(sales_order))
			list_ctrl.SetStringItem(index, 2, '{}'.format(item))
			list_ctrl.SetStringItem(index, 3, '{}'.format(production_order))
			list_ctrl.SetStringItem(index, 4, '{}'.format(material))
			list_ctrl.SetStringItem(index, 5, '{}'.format(sold_to_name))
			list_ctrl.SetStringItem(index, 6, '{}'.format(hours_standard))

			list_ctrl.SetStringItem(index, 7, '{}'.format(date_requested_de_release))
			list_ctrl.SetStringItem(index, 8, '{}'.format(date_planned_de_release))
			list_ctrl.SetStringItem(index, 9, '{}'.format(date_suggested_de_start))

			list_ctrl.SetStringItem(index, 10, '{}'.format(design_engineer))
			list_ctrl.SetStringItem(index, 11, '{}'.format(design_status))
			list_ctrl.SetStringItem(index, 12, '{}'.format(mechanical_status))
			list_ctrl.SetStringItem(index, 13, '{}'.format(electrical_status))
			list_ctrl.SetStringItem(index, 14, '{}'.format(structural_status))
			list_ctrl.SetStringItem(index, 15, '{}'.format(mechanical_engineer))
			list_ctrl.SetStringItem(index, 16, '{}'.format(electrical_engineer))
			list_ctrl.SetStringItem(index, 17, '{}'.format(structural_engineer))
			list_ctrl.SetStringItem(index, 18, '{}'.format(mechanical_cad_designer))
			list_ctrl.SetStringItem(index, 19, '{}'.format(electrical_cad_designer))
			list_ctrl.SetStringItem(index, 20, '{}'.format(structural_cad_designer))

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)
			
			#cap column width at max 400
			if list_ctrl.GetColumnWidth(index) > 400:
				list_ctrl.SetColumnWidth(index, 400)
		
		#hide id column
		list_ctrl.SetColumnWidth(0, 0)
		
		list_ctrl.Thaw()


	def refresh_list_exceptions_de(self, event=None):
		list_ctrl = ctrl(self, 'list:exceptions_de')
		list_ctrl.Freeze()
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

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

				sold_to_name,
				hours_standard,

				date_requested_de_release,
				date_planned_de_release,
				
				DateDiff(Day, date_requested_de_release, date_planned_de_release) AS days_off_by,

				design_engineer,
				design_status,
				mechanical_status,
				electrical_status,
				structural_status,
				mechanical_engineer,
				electrical_engineer,
				structural_engineer,
				mechanical_cad_designer,
				electrical_cad_designer,
				structural_cad_designer
			FROM
				orders.view_systems
			WHERE
				date_actual_de_release IS NULL AND
				production_order IS NOT NULL AND
				--days_off_by > 0 AND
				date_planned_de_release > date_requested_de_release AND
				status <> 'Canceled'
			ORDER BY
				days_off_by DESC
			''')
		
		#insert records into list
		for index, record in enumerate(records):
			#format all fields as strings
			formatted_record = []
			for field in record:
				if field == None:
					field = ''
					
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%Y')
					
				else:
					pass
					
				formatted_record.append(field)

			id, sales_order, item, production_order, material, sold_to_name, hours_standard, \
			date_requested_de_release, date_planned_de_release, days_off_by, \
			design_engineer, design_status, mechanical_status, electrical_status, structural_status, \
			mechanical_engineer, electrical_engineer, structural_engineer, \
			mechanical_cad_designer, electrical_cad_designer, structural_cad_designer = formatted_record
			
			#remove the decimals from the std hours
			try:
				hours_standard = round(hours_standard, 1)
			except:
				pass

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(sales_order))
			list_ctrl.SetStringItem(index, 2, '{}'.format(item))
			list_ctrl.SetStringItem(index, 3, '{}'.format(production_order))
			list_ctrl.SetStringItem(index, 4, '{}'.format(material))
			list_ctrl.SetStringItem(index, 5, '{}'.format(sold_to_name))
			list_ctrl.SetStringItem(index, 6, '{}'.format(hours_standard))

			list_ctrl.SetStringItem(index, 7, '{}'.format(date_requested_de_release))
			list_ctrl.SetStringItem(index, 8, '{}'.format(date_planned_de_release))
			list_ctrl.SetStringItem(index, 9, '{}'.format(days_off_by))

			list_ctrl.SetStringItem(index, 10, '{}'.format(design_engineer))
			list_ctrl.SetStringItem(index, 11, '{}'.format(design_status))
			list_ctrl.SetStringItem(index, 12, '{}'.format(mechanical_status))
			list_ctrl.SetStringItem(index, 13, '{}'.format(electrical_status))
			list_ctrl.SetStringItem(index, 14, '{}'.format(structural_status))
			list_ctrl.SetStringItem(index, 15, '{}'.format(mechanical_engineer))
			list_ctrl.SetStringItem(index, 16, '{}'.format(electrical_engineer))
			list_ctrl.SetStringItem(index, 17, '{}'.format(structural_engineer))
			list_ctrl.SetStringItem(index, 18, '{}'.format(mechanical_cad_designer))
			list_ctrl.SetStringItem(index, 19, '{}'.format(electrical_cad_designer))
			list_ctrl.SetStringItem(index, 20, '{}'.format(structural_cad_designer))

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)
			
			#cap column width at max 400
			if list_ctrl.GetColumnWidth(index) > 400:
				list_ctrl.SetColumnWidth(index, 400)
		
		#hide id column
		list_ctrl.SetColumnWidth(0, 0)
		
		list_ctrl.Thaw()

	def refresh_list_sent_to_mmg(self, event=None):
		list_ctrl = ctrl(self, 'list:sent_to_mmg')
		list_ctrl.Freeze()
		list_ctrl.DeleteAllItems()
		list_ctrl.clean_headers()

		records = db.query('''
			SELECT TOP 30
				orders.root.id,
				orders.root.sales_order,
				orders.root.item,
				orders.root.production_order,
				orders.root.material,
				orders.root.sold_to_name,
				
				dbo.mmg_uploads.filename,
				dbo.mmg_uploads.when_uploaded
			FROM
				orders.root
			RIGHT JOIN
				dbo.mmg_uploads ON orders.root.production_order = dbo.mmg_uploads.production_order
			ORDER BY
				dbo.mmg_uploads.when_uploaded DESC
			''')
		
		#insert records into list
		for index, record in enumerate(records):
			#format all fields as strings
			formatted_record = []
			for field in record:
				if field == None:
					field = ''
					
				elif isinstance(field, dt.datetime):
					field = field.strftime('%m/%d/%y %I:%M %p')
					
				else:
					pass
					
				formatted_record.append(field)

			id, sales_order, item, production_order, material, sold_to_name, \
			filename, when_uploaded = formatted_record

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(id))
			list_ctrl.SetStringItem(index, 1, '{}'.format(sales_order))
			list_ctrl.SetStringItem(index, 2, '{}'.format(item))
			list_ctrl.SetStringItem(index, 3, '{}'.format(production_order))
			list_ctrl.SetStringItem(index, 4, '{}'.format(material))
			list_ctrl.SetStringItem(index, 5, '{}'.format(sold_to_name))

			list_ctrl.SetStringItem(index, 6, '{}'.format(filename))
			list_ctrl.SetStringItem(index, 7, '{}'.format(when_uploaded))

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




class OrdManApp(wx.App):
	def OnInit(self):
		#show splash screen
		gn.splash_frame = AS.AdvancedSplash(None,
			bitmap=wx.Bitmap(gn.resource_path("Splash.png"), wx.BITMAP_TYPE_PNG), 
			timeout=2500, agwStyle=AS.AS_TIMEOUT | AS.AS_CENTER_ON_SCREEN)
			
		wx.Yield()
			
		return True


def handle_gui_exception(exc_type, exc_value, exc_traceback):
	err_msg = '\n'.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
	dlg = wx.MessageDialog(None, err_msg, 'An error occurred!', wx.OK | wx.ICON_ERROR)
	dlg.ShowModal()
	dlg.Destroy()

sys.excepthook = handle_gui_exception


if __name__ == '__main__':
	print 'starting app'

	try:
		gn.app = OrdManApp(False)
		
		try:
			db.eng04_connection = db.connect_to_eng04_database()
		except Exception as ee:
			if ee[0] == 'IM002':
				#try to setup the DSN
				print 'Trying to setup the DSN...'
				call_command = r'''cd C:\WINDOWS\system32 & ODBCConf ConfigDSN "SQL Server" "DSN=eng04_sql|SERVER=KW_SQL_5|DATABASE=eng04_sql|Trusted_Connection=yes"'''
				if subprocess.call(call_command, shell=True):
					wx.MessageBox('Failed to setup ENG04_SQL DSN. Contact IT.\n\n{}'.format(call_command), 'An error occurred!', wx.OK | wx.ICON_ERROR)
				db.eng04_connection = db.connect_to_eng04_database()
		
		if db.eng04_connection == None:
			wx.MessageBox('Could not connect to database ENG04_SQL on server KW_SQL_5.\nGet with IT and establish a connection.', 'Database Connection Problem.', wx.OK | wx.ICON_ERROR)
		
		#load GUI resource
		xrc.XmlResource.Get().Load(gn.resource_path('interface.xrc'))
		
		gn.login_frame = LoginFrame(None)

		gn.app.MainLoop()

	except Exception as e:
		help_message = ''
		try:
			if e[0] == 'IM002':
				help_message = "Your DSN pointing to the ENG04_SQL database is probably not set up correctly.\n\n"

			if e[0] == '28000':
				help_message = "It appears that your username is not authorized to access the ENG04_SQL database\nRequest permission from IT\n\n"
		except Exception as eee:
			print 'eee:', eee

		wx.MessageBox('{}{}\n\n{}'.format(help_message, e, traceback.format_exc()), 'An error occurred!', wx.OK | wx.ICON_ERROR)
		print 'An error occurred!'

	print 'Attempting to end program (outside of app)'

	try: db.eng04_connection.close()
	except Exception as e: print e

	os._exit(0)
