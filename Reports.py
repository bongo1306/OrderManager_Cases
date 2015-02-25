#!/usr/bin/env python
# -*- coding: utf8 -*-

import sys
import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import workdays
import datetime as dt
import General as gn
import Database as db


class ReportsTab(object):
	def init_reports_tab(self):
		#Bindings
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_de_release_stats, id=xrc.XRCID('button:calc_de_release_stats'))
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_de_release_vs_basic_finish, id=xrc.XRCID('button:calc_de_release_vs_basic_finish'))
		self.Bind(wx.EVT_BUTTON, self.on_click_calc_de_eng_errors_vs_actual_finish, id=xrc.XRCID('button:calc_de_eng_errors_vs_actual_finish'))

		self.init_de_release_stats()
		self.init_de_release_vs_basic_finish()
		self.init_de_eng_errors_vs_actual_finish()

	
	def init_de_release_stats(self):
		#default dates to cover last week
		today = dt.date.today()
		first_day_last_week = today - dt.timedelta(days=today.weekday()-2) - dt.timedelta(days=3) #made if default to current week for Paul Juder # - dt.timedelta(weeks=1)
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
				material <> 'SPARTCOLS' AND
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
				#thinking positively, just assume on time if item was never scheduled to begin with.
				ontime_by_request += 1
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
			
			est_de_hours += sum(db.query("SELECT mean_hours FROM orders.material_eto_hour_estimates WHERE material='{}'".format(material)))
			
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
				

		report.AppendText("{} items were released by DE between {} and {}, (excluding SPARTCOLS)\n\n".format(len(rel_items_data), start_date.strftime('%m/%d/%Y'), end_date.strftime('%m/%d/%Y')))
		
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
		
		report.AppendText("\n\nElvis, remember this is still using dbo.std_family_hour_estimates...")
		
		report.ShowPosition(0)



	def init_de_release_vs_basic_finish(self):
		#default to this year
		ctrl(self, 'text:de_release_vs_basic_finish_year').SetValue('{}'.format(dt.date.today().year))
		
		#set up list headers
		list_ctrl = ctrl(self, 'list:de_release_vs_basic_finsh')

		list_ctrl.printer_header = 'DE Release vs Basic Finish'
		list_ctrl.printer_font_size = 8
		
		column_names = ['Week', 'Start Date', 'End Date', 'Basic Finish Std Hours', 'DE Released Std Hours', 'Percent Std Hours']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)


	def init_de_eng_errors_vs_actual_finish(self):
		#default to this year
		ctrl(self, 'text:de_eng_errors_vs_actual_finish_year').SetValue('{}'.format(dt.date.today().year))
		
		#set up list headers
		list_ctrl = ctrl(self, 'list:de_eng_errors_vs_actual_finish')

		list_ctrl.printer_header = 'Engineering Errors vs Actual Finish'
		list_ctrl.printer_font_size = 8
		
		column_names = ['Week', 'Start Date', 'End Date', 'Actual Finish Std Hours', 'Eng Errors', 'Errors per 1000 Std Hours']

		for index, column_name in enumerate(column_names):
			list_ctrl.InsertColumn(index, column_name)


	def on_click_calc_de_release_vs_basic_finish(self, event):
		#determine the starting date of the first week of the year
		year_date = dt.datetime.strptime('{}-1-1'.format(ctrl(self, 'text:de_release_vs_basic_finish_year').GetValue()), '%Y-%m-%d').date()
		base_start_date = year_date - dt.timedelta(days=year_date.weekday()-2) - dt.timedelta(days=3)

		list_ctrl = ctrl(self, 'list:de_release_vs_basic_finsh')
		list_ctrl.DeleteAllItems()
		#list_ctrl.Freeze()
		list_ctrl.clean_headers()

		index = -1
		for week in range(1, 54):
			index += 1
			
			start_date = base_start_date + dt.timedelta(weeks=week-1)
			end_date = start_date + dt.timedelta(days=6)

			records = db.query('''
				SELECT
					hours_standard
				FROM
					orders.view_systems
				WHERE
					status <> 'Canceled' AND
					date_basic_finish >= '{}' AND
					date_basic_finish <= '{}'
			'''.format(start_date, end_date))

			try:
				basic_finish_std_hours = sum(records)
			except:
				basic_finish_std_hours = 0


			records = db.query('''
				SELECT
					hours_standard
				FROM
					orders.view_systems
				WHERE
					status <> 'Canceled' AND
					date_actual_de_release IS NOT NULL AND
					date_basic_finish >= '{}' AND
					date_basic_finish <= '{}'
			'''.format(start_date, end_date))

			try:
				de_release_std_hours = sum(records)
			except:
				de_release_std_hours = 0


			try:
				percent = float(de_release_std_hours/basic_finish_std_hours)*100.
			except:
				percent = 0.0

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(week))
			list_ctrl.SetStringItem(index, 1, '{}'.format(start_date.strftime('%m/%d/%Y')))
			list_ctrl.SetStringItem(index, 2, '{}'.format(end_date.strftime('%m/%d/%Y')))
			list_ctrl.SetStringItem(index, 3, '{:.1f}'.format(basic_finish_std_hours))
			list_ctrl.SetStringItem(index, 4, '{:.1f}'.format(de_release_std_hours))
			list_ctrl.SetStringItem(index, 5, '{:.0f}%'.format(percent))

			#color the row if it's not a 100%
			if percent < 99 and basic_finish_std_hours > 0:
				list_ctrl.SetItemBackgroundColour(index, '#FFE6E6')


			#color the row if it's the current week
			if dt.date.today() >= start_date and dt.date.today() <= end_date:
				list_ctrl.SetItemBackgroundColour(index, '#FFF082')
				
			wx.Yield()

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)

		#list_ctrl.Thaw()


	def on_click_calc_de_eng_errors_vs_actual_finish(self, event):
		#determine the starting date of the first week of the year
		year_date = dt.datetime.strptime('{}-1-1'.format(ctrl(self, 'text:de_eng_errors_vs_actual_finish_year').GetValue()), '%Y-%m-%d').date()
		base_start_date = year_date - dt.timedelta(days=year_date.weekday()-2) - dt.timedelta(days=3)

		list_ctrl = ctrl(self, 'list:de_eng_errors_vs_actual_finish')
		list_ctrl.DeleteAllItems()
		#list_ctrl.Freeze()
		list_ctrl.clean_headers()

		index = -1
		for week in range(1, 54):
			index += 1
			
			start_date = base_start_date + dt.timedelta(weeks=week-1)
			end_date = start_date + dt.timedelta(days=6)

			records = db.query('''
				SELECT
					hours_standard
				FROM
					orders.view_systems
				WHERE
					(date_actual_finish >= '{}' AND date_actual_finish <= '{}') OR
					(date_actual_finish IS NULL AND date_shipped IS NOT NULL AND date_shipped >= '{}' AND date_shipped <= '{}')
			'''.format(start_date, end_date, start_date, end_date))
			
			try:
				actual_finish_std_hours = sum(records)
			except:
				actual_finish_std_hours = 0


			records = db.query('''
				SELECT
					ecrs.id
				FROM
					dbo.ecrs
				INNER JOIN orders.view_systems ON
					(
					CHARINDEX(view_systems.production_order, ecrs.item) > 0 OR 
					CHARINDEX(view_systems.bpcs_item, ecrs.item) > 0
					)
				WHERE
					ecrs.reason = 'Engineering Error' AND
					((view_systems.date_actual_finish >= '{}' AND view_systems.date_actual_finish <= '{}') OR
					(view_systems.date_actual_finish IS NULL AND view_systems.date_shipped IS NOT NULL AND view_systems.date_shipped >= '{}' AND view_systems.date_shipped <= '{}'))
			'''.format(start_date, end_date, start_date, end_date))
			
			try:
				engineering_errors = len(records)
			except:
				engineering_errors = 0

			try:
				engineering_error_rate = engineering_errors / (float(actual_finish_std_hours)/1000.)
			except:
				engineering_error_rate = 0

			list_ctrl.InsertStringItem(sys.maxint, '#')
			list_ctrl.SetStringItem(index, 0, '{}'.format(week))
			list_ctrl.SetStringItem(index, 1, '{}'.format(start_date.strftime('%m/%d/%Y')))
			list_ctrl.SetStringItem(index, 2, '{}'.format(end_date.strftime('%m/%d/%Y')))
			list_ctrl.SetStringItem(index, 3, '{:.1f}'.format(actual_finish_std_hours))
			list_ctrl.SetStringItem(index, 4, '{}'.format(engineering_errors))
			list_ctrl.SetStringItem(index, 5, '{:.1f}'.format(engineering_error_rate))

			#color the row if it's not a 100%
			#if percent < 99 and basic_finish_std_hours > 0:
			#	list_ctrl.SetItemBackgroundColour(index, '#FFE6E6')


			#color the row if it's the current week
			if dt.date.today() >= start_date and dt.date.today() <= end_date:
				list_ctrl.SetItemBackgroundColour(index, '#FFF082')
			
			wx.Yield()

		#auto fit the column widths
		for index in range(list_ctrl.GetColumnCount()):
			list_ctrl.SetColumnWidth(index, wx.LIST_AUTOSIZE_USEHEADER)

		#list_ctrl.Thaw()
