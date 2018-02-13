#!/usr/bin/env python
# -*- coding: utf8 -*-

import os
import sys
import wx
import re #for regular expressions

import datetime as dt
import workdays
import operator


updates_dir = r"\\cbssrvit\kwfiles\Management Software\OrderManager_Case\release"
NightlyBuildPath = r"\\cbssrvit\kwfiles\Management Software\OrderManager_Case\release"

#global variable hold reference to wxGUI
app = None
user = None

#read-only mode
mode_readonly = 'False'

#references needed for when closing on software update
splash_frame = None
login_frame = None
main_frame = None


#used in the resource_path function to prevent CWD changes from messing things up
cwd_at_app_start = os.getcwd()
#now set the CWD to the desktop for user's convinience
os.chdir(os.path.join(os.path.expanduser("~"), "Desktop"))

def resource_path(relative):
	try:
		#return os.path.dirname(sys.executable)
		return os.path.join(sys._MEIPASS, relative)
		
	except:
		#return os.path.join(os.getcwd(), relative)
		return os.path.join(cwd_at_app_start, relative)


def wxdate2pydate(date): 
	assert isinstance(date, wx.DateTime) 
	if date.IsValid(): 
		ymd = map(int, date.FormatISODate().split('-')) 
		return dt.date(*ymd) 
	else: 
		return None

#Date conversion functions
def PY2WX(pyDate):

    if pyDate == None:
        return wx.DefaultDateTime

    try:
       dt = pyDate.date()
       wxDate = wx.DateTime()
       wxDate.ParseDate(str(dt))

       if wxDate.GetYear() == 9999:  return wx.DefaultDateTime
       else: return wxDate

    except ValueError:
       return wx.DefaultDateTime

#calculates the turn around days (date completed - date received) between two given dates in python format
def GetTADays(pyRec, pyComp):

	#convert python date to wx date so we can check date validity
	wx_rec = PY2WX(pyRec)
	wx_comp = PY2WX(pyComp)

	if wx_rec.IsValid() == True and wx_comp.IsValid() == True and wx_rec == wx_comp:
			TADays = "0"
	elif wx_rec.IsValid() == True and wx_comp.IsValid() == True:
			TADays = str(workdays.networkdays(pyRec, pyComp)-1)
	else:
			TADays = str("TBD")

	return TADays


def date_range(start_date, end_date):
	for n in range(int ((end_date - start_date).days)):
		yield start_date + dt.timedelta(n)

#gets a string cleaned up for use in an SQL query
#def clean(string):
#	return string.replace("'", "''")

#replaces non-ascii characters in string for use in an SQL query
# was previously called "deep_clean"...
def clean(string):
	return string.replace(u"¹", "^1").replace(u"²", "^2").replace(u"³", "^3").replace(u"½", "1/2").replace(u"¼", "1/4").replace(u"¾", "3/4") \
		.replace(u"⅛", "1/8").replace(u"⅜", "3/8").replace(u"⅝", "5/8").replace(u"⅞", "7/8").replace(u"⁄", "/").replace(u"°", " degrees") \
		.replace(u"©", "(c)").replace(u"µ", "micro").replace(u"Ω", "ohm").replace(u"®", "(r)").replace(u"™", "(TM)").replace(u"¢", "Cents") \
		.replace(u"€", "Euro").replace(u"“", '''"''').replace(u"”", '''"''').replace(u"‘", "'").replace(u"’", "'").replace(u"′", "'") \
		.replace(u"″", '''"''').replace(u"‐", "-").replace(u"‑", "-").replace(u"–", "-").replace(u"—", "-").replace(u"―", "-").replace(u"…", "...") \
		.replace(u"×", "x").replace(u"÷", "/").replace(u"•", "*").replace(u"", "*").replace(u"'", "''")


def multy_split(s, seps):
	res = [s]
	for sep in seps:
		s, res = res, []
		for seq in s:
			#res += seq.split(sep)
			res += re.split('({})(?i)'.format(sep), seq)
	return res


def humanize_bytes(bytes, precision=1):
	abbrevs = (
		(1<<50L, 'PB'),
		(1<<40L, 'TB'),
		(1<<30L, 'GB'),
		(1<<20L, 'MB'),
		(1<<10L, 'kB'),
		(1, 'bytes'))
		
	if bytes == 1:
		return '1 byte'
		
	for factor, suffix in abbrevs:
		if bytes >= factor:
			break
			
	return '%.*f %s' % (precision, bytes / factor, suffix)

	
def remove_notebook_page(notebook, page_name):
	tabs_to_remove = []
		
	for index in range(notebook.GetPageCount()):
		if notebook.GetPageText(index).strip() == page_name:
			tabs_to_remove.append(index)

	removed_count = 0
	for index in tabs_to_remove:
		notebook.RemovePage(index-removed_count)
		removed_count += 1


def rename_notebook_page(notebook, base_page_name, new_page_name):
	for index in range(notebook.GetPageCount()):
		if base_page_name in notebook.GetPageText(index):
			notebook.SetPageText(index, new_page_name)


