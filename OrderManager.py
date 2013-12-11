#!/usr/bin/env python
# -*- coding: utf8 -*-
version = '0.1'

import sys
import os
import traceback
import subprocess
import threading

import wx				#wxWidgets used as the GUI
from wx import xrc		#allows the loading and access of xrc file (xml) that describes GUI
ctrl = xrc.XRCCTRL		#define a shortined function name (just for convienience)
import wx.lib.agw.advancedsplash as AS
#import wx.lib.inspection

import ConfigParser #for reading local config data (*.cfg)

import datetime as dt

import BetterListCtrl
import General as gn
import Database as db
import Search


def check_for_updates():
	try:
		with open(os.path.join(gn.updates_dir, "updates.txt")) as file:
			lines = file.readlines()
			latest_version = lines[-1].split('`')[0]
			
			if version != latest_version:
				wx.CallAfter(open_software_update_frame)
	except Exception as e:
		print 'Failed update check:', e

def open_software_update_frame():
	SoftwareUpdateFrame(None)


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


	def on_click_login(self, event):
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
			## don't need to store it??? gn.main_frame = MainFrame(None)
			MainFrame(None)
			
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

		#misc
		self.SetTitle('OrderManager v{} - Logged in as {}'.format(version, gn.user))
		#OrderScheduler v{} - Logged in as{} {}'.format(version, self.user.split(',')[-1], self.user.split(',')[0]))

		self.init_search_tab()

		self.Show()
		
		#ItemFrame(self, id=12585385)
		

	def on_close_frame(self, event):
		print 'called on_close_frame'
		self.Destroy()



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

		#misc
		self.SetTitle('Item ID {}'.format(self.id))
		self.SetSize((800, 600))
		self.Center()


		self.Show()
		
		
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
