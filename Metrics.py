import wx
from wx import xrc
ctrl = xrc.XRCCTRL
import General as gn
import datetime
from win32com.client import Dispatch
from wx.html import HtmlEasyPrinting
from operator import itemgetter
import os

class MetricsDialog(wx.Dialog):
	def __init__(self, parent):
		#load XRC description
		pre = wx.PreDialog()
		res = xrc.XmlResource.Get()
		res.LoadOnDialog(pre, parent, "MetricsDialog")
		self.PostCreate(pre)
		ImgPath = str(os.path.split(__file__)[0]) + "\Images\Metrics.png"
		self.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))


		self.SelectedCustomers = []
		self.SearchCursor = None
		self.SearchResultGrid = []
		self.SearchInProgress = False
		self.dialog = None
		self.HeaderSortReversed = False
		self.progress = 0
		self.sql = ""
		SearchIndex = 0                              # zero based index of current search displayed


		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.Bind(wx.EVT_TOOL, self.OnBtnMetricsRun, id=xrc.XRCID('m_BtnMetricsRun'))
		self.Bind(wx.EVT_TOOL, self.OnBtnMetricsSelectAll, id=xrc.XRCID('m_BtnMetricsSelectAll'))
		self.Bind(wx.EVT_TOOL, self.OnBtnMetricsSelectNone, id=xrc.XRCID('m_BtnMetricsSelectNone'))

		self.Bind(wx.EVT_BUTTON, self.OnBtnMetricsPrint, id=xrc.XRCID('m_BtnMetricsPrint'))
		self.Bind(wx.EVT_BUTTON, self.OnBtnMetricsExcel, id=xrc.XRCID('m_BtnMetricsExcel'))

		self.Bind(wx.EVT_LIST_COL_CLICK, self.OnHeaderSort,id=xrc.XRCID('m_ListMetricsResults') )


		self.m_ComboAnalysisType= wx.FindWindowByName('m_ComboAnalysisType')
		self.m_MetDateReceivedStart = wx.FindWindowByName('m_MetDateReceivedStart')
		self.m_MetDateReceivedEnd = wx.FindWindowByName('m_MetDateReceivedEnd')
		self.m_MetDateDueStart = wx.FindWindowByName('m_MetDateDueStart')
		self.m_MetDateDueEnd = wx.FindWindowByName('m_MetDateDueEnd')
		self.m_MetDateCompStart = wx.FindWindowByName('m_MetDateCompStart')
		self.m_MetDateCompEnd = wx.FindWindowByName('m_MetDateCompEnd')
		self.m_ComboMetricsAE = wx.FindWindowByName('m_ComboMetricsAE')
		self.m_ComboMetricsSP = wx.FindWindowByName('m_ComboMetricsSP')
		self.m_CheckListMetCustomers = wx.FindWindowByName('m_CheckListMetCustomers')
		self.m_ListMetricsResults = wx.FindWindowByName('m_ListMetricsResults')

		self.m_CheckListMetCustomers.Bind( wx.EVT_LISTBOX, self.OnCustomerListBoxSelect)


		# ----- initialize the various search controls controls -----------------------
		self.m_ComboAnalysisType.Select(0)

		wxDate = wx.DateTime()
		wxDate.ParseDate("01/01/2000")
		self.m_MetDateReceivedStart.SetValue(wxDate)
		self.m_MetDateDueStart.SetValue(wxDate)
		self.m_MetDateCompStart.SetValue(wxDate)

		wxDate.ParseDate("01/01/2025")
		self.m_MetDateReceivedEnd.SetValue(wxDate)
		self.m_MetDateDueEnd.SetValue(wxDate)
		self.m_MetDateCompEnd.SetValue(wxDate)

		self.Headers = ["Customer Name", "# of Quotes", "# of Orders", "Equip. Quotes ($)",	"Buyouts Quotes ($)",
                        "Total Quotes ($)", "Equip. Orders ($)", "Buyouts Orders ($)",	"Total Orders ($)",
                        "Hit Ratio (%)", "Conversion  Ratio (%)", "Total Order $$ per Quote"]
		self.Footers = ["Total",0,0,0,0,0,0,0,0,0,0,0]

		self.ColumnCount = len(self.Headers)
		for i in range(0, self.ColumnCount):
				self.m_ListMetricsResults.InsertColumn(i, self.Headers[i])
				self.m_ListMetricsResults.SetColumnWidth(i, wx.LIST_AUTOSIZE_USEHEADER)

		self.m_ListMetricsResults.SetColumnWidth(0, 100)


		#--- end initialize search controls -----------------------------

		#misc
		self.SetTitle('Run Metrics')

	def PopulateLists(self):

		self.PopulateAE()
		self.PopulateSP()
		self.PopulateCustomers()

	def PopulateAE(self):

		sql = "SELECT DISTINCT Assigned FROM dbo.QuoteMaker"
		self.SearchCursor.execute(sql)
		row = self.SearchCursor.fetchone()
		AEList = []

		while row != None:
				AE =  str(row.Assigned).strip()
				if len(AE) > 1:AEList.append(AE)
				row = self.SearchCursor.fetchone()


		self.m_ComboMetricsAE.SetItems(AEList)
		self.m_ComboMetricsAE.Append("ALL")
		self.m_ComboMetricsAE.SetStringSelection("ALL")  #Select ALL as the default option

	def PopulateSP(self):

		sql = "SELECT DISTINCT Saleperson FROM dbo.QuoteMaker"
		self.SearchCursor.execute(sql)
		row = self.SearchCursor.fetchone()
		SPList = []

		while row != None:
				SP =  str(row.Saleperson).strip()
				if len(SP) > 1:SPList.append(SP)
				row = self.SearchCursor.fetchone()

		self.m_ComboMetricsSP.SetItems(SPList)
		self.m_ComboMetricsSP.Append("ALL")
		self.m_ComboMetricsSP.SetStringSelection("ALL")   #Select ALL as the default option

	def PopulateCustomers(self):

		sql = "SELECT DISTINCT CustomerKey FROM dbo.QuoteMaker"
		self.SearchCursor.execute(sql)
		row = self.SearchCursor.fetchone()
		custList = []

		while row != None:
				Cust =  str(row.CustomerKey).strip()
				if len(Cust) > 1:custList.append(Cust)
				row = self.SearchCursor.fetchone()

		custList = list(set(custList))
		self.m_CheckListMetCustomers.SetItems(custList)
		self.OnBtnMetricsSelectAll(None)                       #Set all customers as checked

    #Select all customers to be in included in the metrics
	def OnBtnMetricsSelectAll(self, event):
		count = self.m_CheckListMetCustomers.GetCount()
		for i in range(0, count):
				self.m_CheckListMetCustomers.Check(i, True)

    # Select none of the customers to be included in the metrics (useful for subsequent selection of only a few customers)
	def OnBtnMetricsSelectNone(self, event):
		count = self.m_CheckListMetCustomers.GetCount()
		for i in range(0, count):
				self.m_CheckListMetCustomers.Check(i, False)


	def OnCustomerListBoxSelect(self, event):
		self.m_CheckListMetCustomers.SetSelection(-1) #do not want any items shown selected in this box

	def on_close_frame(self, event):
		self.Destroy()

    #############################################################################################################
	def OnBtnMetricsRun(self, event):

		if self.SearchInProgress == True: return  #already searching so ignore the search request
		self.SearchInProgress = True

		self.m_ListMetricsResults.DeleteAllItems()

        # Get a list of customers to be included in the metrics analysis (boxes checked by the user)
		self.SelectedCustomers = self.m_CheckListMetCustomers.GetCheckedStrings()
		NumSelCustomers = len(self.SelectedCustomers)

		if NumSelCustomers == 0:
				self.SearchInProgress = False
				return

        #Headers are appended to the SearchResultGrid so that the grid variable can be passed on to excel or the print
		#function and the headers will be exported as well
		del self.SearchResultGrid[:]
		#self.SearchResultGrid.append(self.Headers)
		self.Footers = ["Total",0,0,0,0,0,0,0,0,0,0,0]

        #Show the progress bar
		msg = "This may take a few minutes, depending on the number of customers, " \
                                 "and the speed of your computer...\n "

		self.dialog = wx.ProgressDialog("Running Metrics", msg,maximum=NumSelCustomers,
                                        style=wx.PD_ELAPSED_TIME|wx.PD_ESTIMATED_TIME|wx.PD_REMAINING_TIME|
                                              wx.PD_APP_MODAL|wx.PD_AUTO_HIDE|wx.PD_SMOOTH|wx.PD_CAN_ABORT)
		self.progress = 1

        # A row will be added to self.SearchResultGrid for each customer
        # The total number of rows in self.SearchResultGrid should be one more than the number of selected customers
		cont = True
		for i in range(0,NumSelCustomers):
				self.GetCustomerMetrics(self.SelectedCustomers[i])
				percentage = int(100 * self.progress/NumSelCustomers)
				msgUpdate = msg + str(percentage) + "%"
				(cont, skip) = self.dialog.Update(self.progress, msgUpdate)
				self.progress = self.progress + 1

				if cont == False:break

		if cont == False:
				wx.MessageBox("The caclualtion was aborted. Only partial results will be available")




        #Insert main results into list  control
		NumRes = len(self.SearchResultGrid)  #should be equal to NumSelCustomers
		for x in range(0, NumRes):

				self.m_ListMetricsResults.InsertStringItem(x,self.SearchResultGrid[x][0])

				for col in range(1, self.ColumnCount):
						self.m_ListMetricsResults.SetStringItem(x,col, self.SearchResultGrid[x][col])

        #Insert the footer (totals) into the list control
		self.FormatFooter()
		self.InsertResultFooter()

        #Make columns the correct size
		self.SizeColumns()

        #Done with running the metrics
		self.dialog.Destroy()
		self.dialog = None
		self.SearchInProgress = False

    #############################################################################################################
   # Format the footer list
	def FormatFooter(self):

		try:
				self.Footers[9] =  '{0:.0f} %'.format(100.0 * self.Footers[2]/self.Footers[1])   #Average hit ratio
				self.Footers[10] = '{0:.0f} %'.format(100.0 * self.Footers[8]/self.Footers[5])   #Average conversion ratio
				self.Footers[11] = '$ {0:,.0f}'.format(self.Footers[8]/self.Footers[1])           #Average order dollars per quote
		except:
				self.Footers[9] =  ""   #Average hit ratio
				self.Footers[10] = ""  #Average conversion ratio
				self.Footers[11] = ""  #Average order dollars per quote


		self.Footers[0] = "Total"
		self.Footers[1] = '{0:.0f}'.format(self.Footers[1])
		self.Footers[2] = '{0:.0f}'.format(self.Footers[2])
		self.Footers[3] = '$ {0:,.0f}'.format(self.Footers[3])
		self.Footers[4] = '$ {0:,.0f}'.format(self.Footers[4])
		self.Footers[5] = '$ {0:,.0f}'.format(self.Footers[5])
		self.Footers[6] = '$ {0:,.0f}'.format(self.Footers[6])
		self.Footers[7] = '$ {0:,.0f}'.format(self.Footers[7])
		self.Footers[8] = '$ {0:,.0f}'.format(self.Footers[8])


	# Inserts a row with aggregate of the search results into the list box
	def InsertResultFooter(self):

		NumRes = len(self.SearchResultGrid)  #should be equal to NumSelCustomers
		self.m_ListMetricsResults.InsertStringItem(NumRes,self.Footers[0])

		for col in range(1, self.ColumnCount):
				self.m_ListMetricsResults.SetStringItem(NumRes,col, self.Footers[col])

		self.m_ListMetricsResults.SetItemBackgroundColour(NumRes, "#e5e5ff")

    ##############################################################################################################

	def SizeColumns(self):
		for col in range(0, self.ColumnCount):self.m_ListMetricsResults.SetColumnWidth(col, wx.LIST_AUTOSIZE_USEHEADER)

     #############################################################################################################

    #Helper function used by GetCustomerMetrics. It returns 0 when an invalid input is supplied
	def GetNum(self, input):

		try:
				if input == None: return 0.0

				str_value = str(input)
				if len(str_value) == 0: return 0.0

				float_value = float(str_value)
				return float_value

		except ValueError:
				return 0.0

    ############################################################################################################
    # This function gets metrics for a single customer. This function adds one row to the self.SearchResultGrid variable
	def GetCustomerMetrics(self, CustName):

		#Get the search parameters
		self.sql = ""
		QuoteString = " AND (ProjectType=\'Quote\')"
		OrderString = " AND (ProjectType=\'Order\')"

        #----------------------------------------------------------------------
        #Get the customer name search field
		self.sql += "SELECT EquipPrice,BuyoutPrice FROM dbo.QuoteMaker WHERE (CustomerKey=?)"

        #----------------------------------------------------------------------
        #Get the revision number search field
		self.sql += " AND (RevLevel=0)"

        #--------------------------------------------------------------------
        #Get the project type search field
		self.sql += QuoteString

       #-------------------------------------------------------------------
        #date received, requested, completed search ranges
		DateCompStart = self.m_MetDateCompStart.GetValue()           #wx.DateTime variable
		DateCompEnd = self.m_MetDateCompEnd.GetValue()               #wx.DateTime variable
		DateReceivedStart = self.m_MetDateReceivedStart.GetValue()   #wx.DateTime variable
		DateReceivedEnd = self.m_MetDateReceivedEnd.GetValue()       #wx.DateTime variable
		DateDueStart = self.m_MetDateDueStart.GetValue()             #wx.DateTime variable
		DateDueEnd = self.m_MetDateDueEnd.GetValue()                 #wx.DateTime variable

		self.sql += " AND (DateReceived BETWEEN \'" + DateReceivedStart.FormatISODate() + "\'"
		self.sql += " AND " + "\'" + DateReceivedEnd.FormatISODate() + "\')"

		self.sql +=  " AND (DateRequest BETWEEN \'" + DateDueStart.FormatISODate() + "\'"
		self.sql += " AND " + "\'" + DateDueEnd.FormatISODate() + "\')"

		if DateCompStart.IsValid() == True and DateCompEnd.IsValid() == True:
				self.sql += " AND (DateComp BETWEEN \'"+ DateCompStart.FormatISODate() + "\'"
				self.sql += " AND " + "\'" + DateCompEnd.FormatISODate() + "\')"

        #----------------------------------------------------------------------
        #Get the analysis type
		AnalysisType = str(self.m_ComboAnalysisType.GetValue())
		AnalysisType = AnalysisType.lower()

		if AnalysisType.find("only equipment") != wx.NOT_FOUND:
				self.sql += " AND (EquipPrice >= 1) AND ((BuyoutPrice <= 1) OR (BuyoutPrice is NULL))"

		elif AnalysisType.find("only buyouts") != wx.NOT_FOUND:
				self.sql += " AND (BuyoutPrice >= 1) AND ((EquipPrice <= 1) OR (EquipPrice is NULL))"

		elif AnalysisType.find("both") != wx.NOT_FOUND:
				self.sql += " AND (EquipPrice >= 1) AND (BuyoutPrice >= 1)"

        #----------------------------------------------------------------------
        #Get the application engineer and saleperson
		AE = self.m_ComboMetricsAE.GetValue()
		if len(AE) > 1 and AE != "ALL":
				self.sql += " AND (Assigned LIKE \'%" + AE + "%\')"

		SP = self.m_ComboMetricsSP.GetValue()
		if len(SP) > 1 and SP != "ALL":
				self.sql += " AND (Saleperson LIKE \'%" + SP + "%\')"

        #------------------------------------------------------------------------
		# Get the "quotes (revision 0)" metrics for this customer
		self.SearchCursor.execute(self.sql, CustName)
		NumQuotes = 0
		EquipUSDQuotes = 0.0
		BuyoutsUSDQuotes = 0.0
		TotalUSDQuotes = 0.0

		row = self.SearchCursor.fetchone()
		while row != None:

				NumQuotes = NumQuotes + 1
				EquipUSDQuotes = EquipUSDQuotes + self.GetNum(row.EquipPrice)
				BuyoutsUSDQuotes = BuyoutsUSDQuotes + self.GetNum(row.BuyoutPrice)
				TotalUSDQuotes = BuyoutsUSDQuotes + EquipUSDQuotes

				row = self.SearchCursor.fetchone()

        #------------------------------------------------------------------------
		# Get the "orders (revision 0)" metrics for this customer
		self.sql = self.sql.replace(QuoteString, OrderString)
		self.SearchCursor.execute(self.sql, CustName)
		NumOrders = 0
		EquipUSDOrders = 0.0
		BuyoutsUSDOrders = 0.0
		TotalUSDOrders = 0.0

		row = self.SearchCursor.fetchone()
		while row != None:

				NumOrders = NumOrders + 1
				EquipUSDOrders = EquipUSDOrders + self.GetNum(row.EquipPrice)
				BuyoutsUSDOrders = BuyoutsUSDOrders + self.GetNum(row.BuyoutPrice)
				TotalUSDOrders = BuyoutsUSDOrders + EquipUSDOrders

				row = self.SearchCursor.fetchone()

        #------------------------------------------------------------------------
        # Get the Hit Ratio, Conversion  Ratio, Total Order $$ per Quote
		if NumQuotes > 0:
				HitRatio = NumOrders * 100.0/NumQuotes
				HitRatio = '{0:.0f} %'.format(HitRatio)
		else:
				HitRatio = ""

		if TotalUSDQuotes > 0:
				ConversionRatio = TotalUSDOrders * 100.0/TotalUSDQuotes
				ConversionRatio =  '{0:.0f} %'.format(ConversionRatio)
		else:
				ConversionRatio = ""

		if NumQuotes > 0:
				OrderUSDPerQuote = TotalUSDOrders/NumQuotes
				OrderUSDPerQuote =  '$ {:,.0f}'.format(OrderUSDPerQuote)
		else:
				OrderUSDPerQuote = ""

        #Footer has aggregate information for all customers.
		self.Footers[1] += NumQuotes
		self.Footers[2] += NumOrders
		self.Footers[3] += EquipUSDQuotes
 		self.Footers[4] += BuyoutsUSDQuotes
 		self.Footers[5] += TotalUSDQuotes
		self.Footers[6] += EquipUSDOrders
 		self.Footers[7] += BuyoutsUSDOrders
 		self.Footers[8] += TotalUSDOrders


        #convert output to string
		EquipUSDQuotes = '$ {0:,.0f}'.format(EquipUSDQuotes)
		BuyoutsUSDQuotes = '$ {:,.0f}'.format(BuyoutsUSDQuotes)
		TotalUSDQuotes = '$ {:,.0f}'.format(TotalUSDQuotes)

		EquipUSDOrders = '$ {:,.0f}'.format(EquipUSDOrders)
		BuyoutsUSDOrders = '$ {:,.0f}'.format(BuyoutsUSDOrders)
		TotalUSDOrders = '$ {:,.0f}'.format(TotalUSDOrders)


		self.SearchResultGrid.append([CustName, str(NumQuotes), str(NumOrders),EquipUSDQuotes,
                     BuyoutsUSDQuotes, TotalUSDQuotes, EquipUSDOrders,
                     BuyoutsUSDOrders, TotalUSDOrders, HitRatio, ConversionRatio, OrderUSDPerQuote])


   # Sort the results according to the column header clicked by the user
	def OnHeaderSort(self, event):
		column = event.GetColumn()
		self.SearchResultGrid.sort(key=lambda k: self.ConvColFormat(k[column]), reverse=self.HeaderSortReversed)
		self.HeaderSortReversed = not (self.HeaderSortReversed)
		self.m_ListMetricsResults.DeleteAllItems()

       #Insert results into list  control
		NumRes = len(self.SearchResultGrid)  #should be equal to NumSelCustomers
		for x in range(0, NumRes):

				self.m_ListMetricsResults.InsertStringItem(x,self.SearchResultGrid[x][0])

				for col in range(1, self.ColumnCount):
						self.m_ListMetricsResults.SetStringItem(x,col, self.SearchResultGrid[x][col])

        #Insert the self.Footers row containing aggregae customer information
		self.InsertResultFooter()

        #Make columns the correct size
		self.SizeColumns()

    # try converting column format from string to float (for list sorting purposes)
    # if not succcessful (such as for customer name column), then return the original string back
	def ConvColFormat(self, str_val):
		try:
			s = str_val
			if s == None or len(s) == 0: return 0.0
			s = s.replace('$','')
			s = s.replace('%','')
			s = s.replace(',','')
			return float(s)

		except:
			return str_val


    #takes the search records and dumps output into excel
	def OnBtnMetricsExcel(self, event):

		excel = Dispatch('Excel.Application')
		excel.Visible = True
		wb = excel.Workbooks.Add()

		wb.ActiveSheet.Cells(1, 1).Value = 'Transferring data to Excel...'
		wb.ActiveSheet.Columns(1).AutoFit()

		R = len(self.SearchResultGrid)
		C = self.ColumnCount

        #Write the header
		excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(1,C))
		excel_range.Value = self.Headers

		#Write the main results
		excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(2, 1), wb.ActiveSheet.Cells(R+1,C))
		excel_range.Value = self.SearchResultGrid

        #Write the footers (totals)
		excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(R+2, 1),wb.ActiveSheet.Cells(R+2,C))
		excel_range.Value = self.Footers

        #Autofit the columns
		excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(R+2,C))
		excel_range.Columns.AutoFit()


	def OnBtnMetricsPrint(self, event):

		html = '''<table border="1" cellspacing="0">\n'''
		R = len(self.SearchResultGrid)
		C = self.ColumnCount

        #write out the header row
		html += '<tr>'
		for col in range(0,C):
				html += '''<td align="left" valign="top" nowrap>{}&nbsp;</td>\n'''.format(self.Headers[col])
		html += '</tr>'


		#write out data
		for row in range(0, R):

			html += '<tr>'
			for col in range(0,C):
				html += '''<td align="left" valign="top" nowrap>{}&nbsp;</td>\n'''.format(self.SearchResultGrid[row][col])
			html += '</tr>'


        #write out the footer row
		html += '<tr>'
		for col in range(0,C):
				html += '''<td align="left" valign="top" nowrap>{}&nbsp;</td>\n'''.format(self.Footers[col])
		html += '</tr>'

		html += '</table>'

		printer = HtmlEasyPrinting()

		#override these settings if you want
		printer_header = 'Printed on @DATE@, Page @PAGENUM@ of @PAGESCNT@'
		printer_font_size = 8
		printer_paper_type = wx.PAPER_11X17         #wx.PAPER_LETTER
		printer_paper_orientation = wx.LANDSCAPE	#wx.PORTRAIT
		printer_paper_margins = (1, 1, 1, 1)	    #specify like (0, 0, 0, 0), otherwise None means whatever default is

		printer.SetHeader(printer_header)
		printer.SetStandardFonts(printer_font_size)
		printer.GetPrintData().SetPaperId(printer_paper_type)
		printer.GetPrintData().SetOrientation(printer_paper_orientation)

		if printer_paper_margins:
			printer.GetPageSetupData().SetMarginTopLeft((printer_paper_margins[0], printer_paper_margins[1]))
			printer.GetPageSetupData().SetMarginBottomRight((printer_paper_margins[2], printer_paper_margins[3]))

		printer.PrintText(html)
















