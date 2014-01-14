#######################################################################
# snapshotPrinter.py
#
# Created: 12/26/2007 by mld
#
# Description: Displays screenshot image using html and then allows
#			  the user to print it.
#######################################################################
 
import os
import wx
import sys
from wx.html import HtmlEasyPrinting, HtmlWindow

import General as gn


def print_window(frame):
	take_screenshot(frame)
	printer = SnapshotPrinter()
	printer.sendToPrinter()

def take_screenshot(frame):
	""" Takes a screenshot of the screen at give pos & size (rect). """
	print 'Taking screenshot...'
	rect = frame.GetRect()
	# see http://aspn.activestate.com/ASPN/Mail/Message/wxpython-users/3575899
	# created by Andrea Gavana

	# adjust widths for Linux (figured out by John Torres 
	# http://article.gmane.org/gmane.comp.python.wxpython/67327)
	if sys.platform == 'linux2':
		client_x, client_y = frame.ClientToScreen((0, 0))
		border_width = client_x - rect.x
		title_bar_height = client_y - rect.y
		rect.width += (border_width * 2)
		rect.height += title_bar_height + border_width
	
	#limit max screenshot width and height to prevent stupid print bug
	if rect.width > 1000:
		rect.width = 1000
	if rect.height > 700:
		rect.height = 700

	#Create a DC for the whole screen area
	dcScreen = wx.ScreenDC()

	#Create a Bitmap that will hold the screenshot image later on
	#Note that the Bitmap must have a size big enough to hold the screenshot
	#-1 means using the current default colour depth
	bmp = wx.EmptyBitmap(rect.width, rect.height)

	#Create a memory DC that will be used for actually taking the screenshot
	memDC = wx.MemoryDC()

	#Tell the memory DC to use our Bitmap
	#all drawing action on the memory DC will go to the Bitmap now
	memDC.SelectObject(bmp)

	#Blit (in this case copy) the actual screen on the memory DC
	#and thus the Bitmap
	memDC.Blit( 0, #Copy to this X coordinate
				0, #Copy to this Y coordinate
				rect.width, #Copy this width
				rect.height, #Copy this height
				dcScreen, #From where do we copy?
				rect.x, #What's the X offset in the original DC?
				rect.y  #What's the Y offset in the original DC?
				)

	#Select the Bitmap out of the memory DC by selecting a new
	#uninitialized Bitmap
	memDC.SelectObject(wx.NullBitmap)

	img = bmp.ConvertToImage()
	
	'''
	print 'img.GetWidth():', img.GetWidth()
	PhotoMaxSize = 800
	W = img.GetWidth()
	H = img.GetHeight()
	if W > H:
		NewW = PhotoMaxSize
		NewH = PhotoMaxSize * H / W
	else:
		NewH = PhotoMaxSize
		NewW = PhotoMaxSize * W / H
	
	img = img.Scale(NewW, NewH)
	#img = img.Scale(500, 500) 
	'''
	
	fileName = "screenshot.png"
	#img.SaveFile(fileName, wx.BITMAP_TYPE_PNG)
	img.SaveFile(gn.resource_path(fileName), wx.BITMAP_TYPE_PNG)
	
	print '...saving as png!'



class SnapshotPrinter(wx.Frame):

	#----------------------------------------------------------------------
	def __init__(self, title='Snapshot Printer'):
		wx.Frame.__init__(self, None, wx.ID_ANY, title, size=(650,400))

		self.panel = wx.Panel(self, wx.ID_ANY)
		self.printer = HtmlEasyPrinting(name='Printing', parentWindow=None)

		self.html = HtmlWindow(self.panel)
		self.html.SetRelatedFrame(self, self.GetTitle())

		#if not os.path.exists('screenshot.htm'):
		if not os.path.exists(gn.resource_path('screenshot.htm')):
			self.createHtml()
		#self.createHtml() # get rid of!!1 @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
		#self.html.LoadPage('screenshot.htm')
		self.html.LoadPage(gn.resource_path('screenshot.htm'))

		pageSetupBtn = wx.Button(self.panel, wx.ID_ANY, 'Page Setup')
		printBtn = wx.Button(self.panel, wx.ID_ANY, 'Print')
		cancelBtn = wx.Button(self.panel, wx.ID_ANY, 'Cancel')

		self.Bind(wx.EVT_BUTTON, self.onSetup, pageSetupBtn)
		self.Bind(wx.EVT_BUTTON, self.onPrint, printBtn)
		self.Bind(wx.EVT_BUTTON, self.onCancel, cancelBtn)

		sizer = wx.BoxSizer(wx.VERTICAL)
		btnSizer = wx.BoxSizer(wx.HORIZONTAL)

		sizer.Add(self.html, 1, wx.GROW)
		btnSizer.Add(pageSetupBtn, 0, wx.ALL, 5)
		btnSizer.Add(printBtn, 0, wx.ALL, 5)
		btnSizer.Add(cancelBtn, 0, wx.ALL, 5)
		sizer.Add(btnSizer)

		self.panel.SetSizer(sizer)
		self.panel.SetAutoLayout(True)

	#----------------------------------------------------------------------
	def createHtml(self):
		'''
		Creates an html file in the home directory of the application
		that contains the information to display the snapshot
		'''
		print 'creating html...'

		#html = '<html>\n<body>\n<center><img src=myImage.png width=516 height=314></center>\n</body>\n</html>'
		html = "<html>\n<body>\n<center><img src=screenshot.png></center>\n</body>\n</html>"
		
		#f = file('screenshot.htm', 'w')
		f = file(gn.resource_path('screenshot.htm'), 'w')
		f.write(html)
		f.close()

	#----------------------------------------------------------------------
	def onSetup(self, event):
		self.printer.PageSetup()

	#----------------------------------------------------------------------
	def onPrint(self, event):
		self.sendToPrinter()

	#----------------------------------------------------------------------
	def sendToPrinter(self):
		""""""
		self.printer.GetPrintData().SetPaperId(wx.PAPER_LETTER)
		self.printer.GetPrintData().SetOrientation(wx.LANDSCAPE)
		self.printer.GetPageSetupData().SetMarginTopLeft((0, 15))
		self.printer.GetPageSetupData().SetMarginBottomRight((0, 0))
		self.printer.PrintFile(self.html.GetOpenedPage())

	#----------------------------------------------------------------------
	def onCancel(self, event):
		self.Close()

 
class wxHTML(HtmlWindow):
	#----------------------------------------------------------------------
	def __init__(self, parent, id):
		html.HtmlWindow.__init__(self, parent, id, style=wx.NO_FULL_REPAINT_ON_RESIZE)

 
if __name__ == '__main__':
	app = wx.App(False)
	frame = SnapshotPrinter()
	frame.Show()
	app.MainLoop()
