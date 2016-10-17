import wx
from wx import xrc
ctrl = xrc.XRCCTRL
import General as gn
import datetime
import workdays
from win32com.client import Dispatch

from wx.html import HtmlEasyPrinting
import os

class QuoteSearchDialog(wx.Dialog):
        def __init__(self, parent):
                #load XRC description
                pre = wx.PreDialog()
                res = xrc.XmlResource.Get()
                res.LoadOnDialog(pre, parent, "QuoteSearchDialog")
                self.PostCreate(pre)
                ImgPath = str(os.path.split(__file__)[0]) + "\Images\Search-icon.png"
                self.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))

                self.SearchCursor = None
                self.SearchResultGrid = []
                self.SearchInProgress = False
                self.MAX_SEARCH_RESULTS = 5000
                self.sql = ""
                self.OrderCriteria = " ORDER BY DateReceived DESC"
                SearchIndex = 0                              # zero based index of current search displayed
                self.HeaderSortDESC = True


                self.m_ComboSearchProjectType = wx.FindWindowByName('m_ComboSearchProjectType')
                self.m_TextSearchQuoteNum = wx.FindWindowByName('m_TextSearchQuoteNum')
                self.m_ComboSearchRev = wx.FindWindowByName('m_ComboSearchRev')

                self.m_TextSearchSalesOrder = wx.FindWindowByName('m_TextSearchSalesOrder')
                self.m_TextSearchDropShip = wx.FindWindowByName('m_TextSearchDropShip')

                self.m_DateReceivedStart = wx.FindWindowByName('m_DateReceivedStart')
                self.m_DateReceivedEnd = wx.FindWindowByName('m_DateReceivedEnd')
                self.m_DateDueStart = wx.FindWindowByName('m_DateDueStart')
                self.m_DateDueEnd = wx.FindWindowByName('m_DateDueEnd')
                self.m_DateCompStart = wx.FindWindowByName('m_DateCompStart')
                self.m_DateCompEnd = wx.FindWindowByName('m_DateCompEnd')

                self.m_ComboSearchAE = wx.FindWindowByName('m_ComboSearchAE')
                self.m_ComboSearchSP = wx.FindWindowByName('m_ComboSearchSP')

                self.m_TextSearchEquipMIN = wx.FindWindowByName('m_TextSearchEquipMIN')
                self.m_TextSearchEquipMAX = wx.FindWindowByName('m_TextSearchEquipMAX')
                self.m_TextSearchBuyoutMIN = wx.FindWindowByName('m_TextSearchBuyoutMIN')
                self.m_TextSearchBuyoutMAX = wx.FindWindowByName('m_TextSearchBuyoutMAX')

                self.m_TextSearchCustName = wx.FindWindowByName('m_TextSearchCustName')
                self.m_TextSearchCustNum = wx.FindWindowByName('m_TextSearchCustNum')
                self.m_TextSearchShipTO = wx.FindWindowByName('m_TextSearchShipTO')
                self.m_TextSearchNotes = wx.FindWindowByName('m_TextSearchNotes')
                self.m_CheckSearchCMAT = wx.FindWindowByName('m_CheckSearchCMAT')
                self.m_ComboSearchCMAT = wx.FindWindowByName('m_ComboSearchCMAT')
                #self.m_TextQuery = wx.FindWindowByName('m_TextQuery')
                self.m_ListSearchResults = wx.FindWindowByName('m_ListSearchResults')

                #bindings
                self.Bind(wx.EVT_CLOSE, self.on_close_frame)
                self.Bind(wx.EVT_BUTTON, self.OnBtnSearch, id=xrc.XRCID('m_BtnSearch'))
                self.Bind(wx.EVT_BUTTON, self.OnBtnSearchPrint, id=xrc.XRCID('m_BtnSearchPrint'))
                self.Bind(wx.EVT_BUTTON, self.OnBtnSearchExcel, id=xrc.XRCID('m_BtnSearchExcel'))
                self.Bind(wx.EVT_LIST_COL_CLICK, self.OnHeaderSort,id=xrc.XRCID('m_ListSearchResults') )
                self.Bind(wx.EVT_CHECKBOX, self.OnCheckSearchCMAT, id=xrc.XRCID('m_CheckSearchCMAT'))

                # ----- initialize the various search controls controls -----------------------
                self.m_ComboSearchRev.SetStringSelection('ALL')
                self.m_ComboSearchProjectType.SetStringSelection('ALL')

                wxDate = wx.DateTime()
                wxDate.ParseDate("01/01/2000")
                self.m_DateReceivedStart.SetValue(wxDate)
                self.m_DateDueStart.SetValue(wxDate)
                self.m_DateCompStart.SetValue(wx.DefaultDateTime)

                wxDate.ParseDate("01/01/2030")
                self.m_DateReceivedEnd.SetValue(wxDate)
                self.m_DateDueEnd.SetValue(wxDate)
                self.m_DateCompEnd.SetValue(wx.DefaultDateTime)

                self.HeaderDBNames = ["", "RecordKey", "ProjectType", "QuoteNumber", "RevLevel", "SalesOrderNum","DropShipOrderNum",
                                                          "DateReceived" ,"DateRequest","DateComp","", "Assigned","Saleperson","EquipPrice", "BuyoutPrice",
                                                          "CustomerKey","ShipTO","CMAT"]

                self.Headers = ["No.","Record Key" ,"Project Type","Quote Number","Rev","Sales Order No.","Drop Ship No.", "Date Received",
                                   "Date Expected","Date Completed","TA Days", "Application Engineer","Salesperson","Equipment (USD)","Buyouts (USD)",
                                   "Customer Key","Ship TO","CMAT"]

                self.ColumnCount = len(self.Headers)
                for i in range(0, self.ColumnCount):
                                self.m_ListSearchResults.InsertColumn(i, self.Headers[i])
                                self.m_ListSearchResults.SetColumnWidth(i, wx.LIST_AUTOSIZE_USEHEADER)

                self.m_ListSearchResults.SetColumnWidth(0, 50)
                self.m_ListSearchResults.Bind( wx.EVT_LIST_ITEM_ACTIVATED, self.OnClickSearchRes)


                #--- end initialize search controls -----------------------------

                #misc
                self.SetTitle('Quote Search')

        def PopulateAE(self, AEList):

                self.m_ComboSearchAE.SetItems(AEList)
                self.m_ComboSearchAE.Append("ALL")
                self.m_ComboSearchAE.SetStringSelection("ALL")  #Select ALL as the default option

        def PopulateSP(self, SPList):

                self.m_ComboSearchSP.SetItems(SPList)
                self.m_ComboSearchSP.Append("ALL")
                self.m_ComboSearchSP.SetStringSelection("ALL")   #Select ALL as the default option

        def PopulateCMAT(self, CMATList):
                self.m_ComboSearchCMAT.SetItems(CMATList)
                self.m_ComboSearchCMAT.SetStringSelection("ANY")   #Select ANY as the default option

        def on_close_frame(self, event):
                self.Destroy()

        def py2wxdate(date):
                assert isinstance(date, (datetime.datetime, datetime.date))
                tt = date.timetuple()
                dmy = (tt[2], tt[1]-1, tt[0])
                return wx.DateTimeFromDMY(*dmy)

        def wx2pydate(date):
                assert isinstance(date, wx.DateTime)
                if date.IsValid():
                                ymd = map(int, date.FormatISODate().split('-'))
                                return datetime.date(*ymd)
                else:
                                return None

        def GetNumber(self, str_value):

                if len(str_value) == 0:
                                return -1

                try:
                        float_value = float(str_value)
                        return float_value
                except ValueError:
                                return -1

                #takes in a python date and returns its string
        def GetDateString(self,d):
                try:
                        dt_str = str(d.date())
                        return dt_str
                except:
                        return ""

        def OnBtnSearch(self, event):

                if self.SearchInProgress == True:  #already searching so ignore the search request
                                return

                self.SearchInProgress = True
                self.m_ListSearchResults.DeleteAllItems()


                #Get the search parameters
                self.sql = ""
                parameters = []
           #-------------------------------------------------------------------
                #date received, requested, completed search ranges
                DateCompStart = self.m_DateCompStart.GetValue()           #wx.DateTime variable
                DateCompEnd = self.m_DateCompEnd.GetValue()               #wx.DateTime variable
                DateReceivedStart = self.m_DateReceivedStart.GetValue()   #wx.DateTime variable
                DateReceivedEnd = self.m_DateReceivedEnd.GetValue()       #wx.DateTime variable
                DateDueStart = self.m_DateDueStart.GetValue()             #wx.DateTime variable
                DateDueEnd = self.m_DateDueEnd.GetValue()                 #wx.DateTime variable

                self.sql += "SELECT * FROM dbo.QuoteMaker WHERE (DateReceived BETWEEN \'" + DateReceivedStart.FormatISODate() + "\'"
                self.sql += " AND " + "\'" + DateReceivedEnd.FormatISODate() + "\')"

                self.sql +=  " AND (DateRequest BETWEEN \'" + DateDueStart.FormatISODate() + "\'"
                self.sql += " AND " + "\'" + DateDueEnd.FormatISODate() + "\')"

                if DateCompStart.IsValid() == True and DateCompEnd.IsValid() == True:
                                self.sql += " AND (DateComp BETWEEN \'"+ DateCompStart.FormatISODate() + "\'"
                                self.sql += " AND " + "\'" + DateCompEnd.FormatISODate() + "\')"

                #--------------------------------------------------------------------
                #Get the project type search field
                PT = self.m_ComboSearchProjectType.GetValue()
                if len(PT) > 1 and PT != "ALL":
                                self.sql += " AND (ProjectType=\'" + PT + "\')"

                #--------------------------------------------------------------------
                #Get the quote number search field
                QuoteNumber = self.m_TextSearchQuoteNum.GetValue()
                if QuoteNumber.isdigit() == True and len(QuoteNumber) > 1:
                                self.sql += " AND (QuoteNumber=" + QuoteNumber + ")"

                #----------------------------------------------------------------------
                #Get the revision number search field
                Rev = self.m_ComboSearchRev.GetValue()
                if Rev.isdigit() == True:
                                self.sql += " AND (RevLevel=" + Rev + ")"

                #----------------------------------------------------------------------
                #Get the sales and drop ship number search field
                SalesOrder = self.m_TextSearchSalesOrder.GetValue()
                if len(SalesOrder) > 1:
                                SalesOrder = "%" + SalesOrder + "%"
                                self.sql += " AND (SalesOrderNum LIKE ?)"
                                parameters.append(SalesOrder)

                DropShip = self.m_TextSearchDropShip.GetValue()
                if len(DropShip) > 1:
                                DropShip = "%" + DropShip + "%"
                                self.sql += " AND (DropShipOrderNum LIKE ?)"
                                parameters.append(DropShip)

                #----------------------------------------------------------------------
                #Get the application engineer and saleperson
                AE = self.m_ComboSearchAE.GetValue()
                if len(AE) > 1 and AE != "ALL":
                                AE = "%" + AE + "%"
                                self.sql += " AND (Assigned LIKE ?)"
                                parameters.append(AE)

                SP = self.m_ComboSearchSP.GetValue()
                if len(SP) > 1 and SP != "ALL":
                                SP = "%" + SP + "%"
                                self.sql += " AND (Saleperson LIKE ?)"
                                parameters.append(SP)

                #----------------------------------------------------------------------
                #Get the equipment and buyout USD Ranges
                EquipMin = self.m_TextSearchEquipMIN.GetValue()
                EquipMax = self.m_TextSearchEquipMAX.GetValue()
                BOMin = self.m_TextSearchBuyoutMIN.GetValue()
                BOMax = self.m_TextSearchBuyoutMAX.GetValue()

                if self.GetNumber(EquipMin) >= 0:
                                self.sql += " AND (EquipPrice >= " + EquipMin + ")"
                if self.GetNumber(EquipMax) >= 0:
                                self.sql += " AND (EquipPrice <= " + EquipMax + ")"
                if self.GetNumber(BOMin) >= 0:
                                self.sql += " AND (BuyoutPrice >= " + BOMin + ")"
                if self.GetNumber(BOMax) >= 0:
                                self.sql += " AND (BuyoutPrice <= " + BOMax + ")"

                #----------------------------------------------------------------------
                #Get the customer name search field
                CustName = self.m_TextSearchCustName.GetValue()
                if len(CustName) > 1:
                                CustName = "%" + CustName + "%"
                                self.sql += " AND (Customer LIKE ?)"
                                parameters.append(CustName)

                #----------------------------------------------------------------------
                #Get the customer number search field
                CustNum = self.m_TextSearchCustNum.GetValue()
                if len(CustNum) > 1:
                                self.sql += " AND (CustomerNumber=" + CustNum + ")"

                 #-----------------------------------------------------------------------
                #Get the ship to
                ShipAddress = self.m_TextSearchShipTO.GetValue()
                if len(ShipAddress) > 1:
                                ShipAddress = "%" + ShipAddress + "%"
                                self.sql += " AND (ShipTO LIKE ?)"
                                parameters.append(ShipAddress)

                 #-----------------------------------------------------------------------
                #Get the notes
                Notes = self.m_TextSearchNotes.GetValue()
                if len(Notes) > 1:
                                Notes = "%" + Notes + "%"
                                self.sql += " AND (Comments LIKE ?)"
                                parameters.append(Notes)

                #---------------------------------------------------------------------------
                #Get the CMAT if any
                bCheck = self.m_CheckSearchCMAT.GetValue()
                CM = str(self.m_ComboSearchCMAT.GetValue())
                CM = CM.strip()
                if bCheck == True and len(CM) > 1 and CM.find("ANY") == -1:
                        self.sql += " AND (CMAT=\'" + CM + "\')"

                elif bCheck == True and CM.find("ANY") >= 0:
                        self.sql += "AND (CMAT IS NOT NULL)"

                self.sql += self.OrderCriteria
                self.SearchCursor.execute(self.sql, parameters)

                del self.SearchResultGrid[:]

                SearchRecordKeys = []
                row = self.SearchCursor.fetchone()
                NumRes = 0

                while row != None:

                                GridRow = [str(NumRes), str(row.RecordKey),str(row.ProjectType),str(row.QuoteNumber),str(row.RevLevel),
                                                   str(row.SalesOrderNum), str(row.DropShipOrderNum),self.GetDateString(row.DateReceived),
                                                   self.GetDateString(row.DateRequest),self.GetDateString(row.DateComp),
                                                   gn.GetTADays(row.DateReceived, row.DateComp), str(row.Assigned), str(row.Saleperson),
                                                   str(row.EquipPrice),str(row.BuyoutPrice),str(row.CustomerKey), str(row.ShipTO).strip(), str(row.CMAT).strip()]

                                self.SearchResultGrid.append(GridRow)
                                SearchRecordKeys.append(row.RecordKey)

                                row = self.SearchCursor.fetchone()
                                NumRes = NumRes + 1


                Title = "Quote Search: " + str(NumRes) + " results found"
                self.SetTitle(Title)

                if NumRes == 0:   #No search results found
                                self.SearchInProgress = False
                                return

                #Get the parent class DBRecordKeys and replace all keys with just the search keys so we can browse through
                #the search records in the main window
                self.GetParent().DBRecordKeys = SearchRecordKeys
                self.GetParent().Index = 0
                self.GetParent().LoadRecord()

                #Insert results into list  control
                x_max = min(NumRes, self.MAX_SEARCH_RESULTS)
                for x in range(0, x_max):

                        self.m_ListSearchResults.InsertStringItem(x,self.SearchResultGrid[x][0])

                        for col in range(1, self.ColumnCount):
                                self.m_ListSearchResults.SetStringItem(x,col, self.SearchResultGrid[x][col])

                        self.RowColor(x)


                #Make columns the correct size
                self.SizeColumns()

                #End of search
                self.SearchInProgress = False

                #Too many search results found
                msg = "The search produced large number of results. Please consider narrowing the search critera. " \
                          "Displaying first " + str(self.MAX_SEARCH_RESULTS) + " results. You can browse full search results" \
                                                                                                                                   " by returning to the main application window"
                if NumRes > self.MAX_SEARCH_RESULTS: wx.MessageBox(msg,"Search Results", wx.OK | wx.ICON_EXCLAMATION)


        def SizeColumns(self):
                for col in range(0, self.ColumnCount):self.m_ListSearchResults.SetColumnWidth(col, wx.LIST_AUTOSIZE_USEHEADER)


        def RowColor(self, row):
                if row% 2:self.m_ListSearchResults.SetItemBackgroundColour(row, "#e5e5ff")
                else: self.m_ListSearchResults.SetItemBackgroundColour(row, "#ffffe5")

   # Sort the results according to the column header clicked by the user
        def OnHeaderSort(self, event):
                column = event.GetColumn()   # zero based column number

                #Cannot sort by columns that contain information not in the database
                if len (self.HeaderDBNames[column]) == 0: return

                self.OrderCriteria = " ORDER BY " + self.HeaderDBNames[column]
                if self.HeaderSortDESC == True: self.OrderCriteria += " DESC"
                else: self.OrderCriteria += " ASC"
                self.HeaderSortDESC = not (self.HeaderSortDESC)

                self.OnBtnSearch(None)

        # try converting column format from string to float (for list sorting purposes)
        # if not succcessful (such as for customer name column), then return the original string back
        def ConvColFormat(self, str_val):
                try:
                        s = str_val
                        if s == None or len(s) == 0: return 0.0
                        s = s.replace('$','')
                        s = s.replace('%','')
                        return float(s)

                except:
                        return str_val


        def OnCheckSearchCMAT(self, event):
                bCheck = self.m_CheckSearchCMAT.GetValue()
                self.m_ComboSearchCMAT.Enable(bCheck)

        def OnClickSearchRes(self, event):
                item = self.m_ListSearchResults.GetNextSelected(-1)
                RecordKey = self.m_ListSearchResults.GetItemText(item, 1)

                #find the record key in the list
                i = self.GetParent().DBRecordKeys.index(RecordKey)

                #jump to this record on the main window
                if i < 0:
                                return
                else:
                                self.GetParent().Index = i
                                self.GetParent().LoadRecord()
                                self.Destroy()

        def OnBtnSearchPrint(self, event):

                html = '''<table border="1" cellspacing="0">\n'''
                R = min(len(self.SearchResultGrid), self.MAX_SEARCH_RESULTS)
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

                html += '</table>'

                printer = HtmlEasyPrinting()

                #override these settings if you want
                printer_header = 'Printed on @DATE@, Page @PAGENUM@ of @PAGESCNT@'
                printer_font_size = 8
                printer_paper_type = wx.PAPER_11X17         #wx.PAPER_LETTER
                printer_paper_orientation = wx.LANDSCAPE        #wx.PORTRAIT
                printer_paper_margins = (1, 1, 1, 1)        #specify like (0, 0, 0, 0), otherwise None means whatever default is

                printer.SetHeader(printer_header)
                printer.SetStandardFonts(printer_font_size)
                printer.GetPrintData().SetPaperId(printer_paper_type)
                printer.GetPrintData().SetOrientation(printer_paper_orientation)

                if printer_paper_margins:
                        printer.GetPageSetupData().SetMarginTopLeft((printer_paper_margins[0], printer_paper_margins[1]))
                        printer.GetPageSetupData().SetMarginBottomRight((printer_paper_margins[2], printer_paper_margins[3]))

                printer.PrintText(html)


        #takes the search records and dumps output into excel
        def OnBtnSearchExcel(self, event):
                
                excel = Dispatch('Excel.Application')
                excel.Visible = True
                wb = excel.Workbooks.Add()
                

                wb.ActiveSheet.Cells(1, 1).Value = 'Transferring data to Excel...'
                wb.ActiveSheet.Columns(1).AutoFit()
                
                R = min(len(self.SearchResultGrid),self.MAX_SEARCH_RESULTS)
                C = self.ColumnCount

                #Write the header
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(1,C))
                excel_range.Value = self.Headers
                
                #specify the excel range that the main data covers
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(2, 1), wb.ActiveSheet.Cells(R+1,C))
                excel_range.Value = self.SearchResultGrid

                #Autofit the columns
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(R+1,C))
                excel_range.Columns.AutoFit()
                


#####################################################################################################################
#####################################################################################################################

class DailyScheduleFrame(wx.Frame):
        def __init__(self, parent):
                #load frame XRC description
                pre = wx.PreFrame()
                res = xrc.XmlResource.Get()
                res.LoadOnFrame(pre, parent, "DailyScheduleFrame")
                self.PostCreate(pre)
                ImgPath = str(os.path.split(__file__)[0]) + "\Images\DailySchedule.png"
                self.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))
                self.SetTitle('Daily Schedule')

                self.m_ListDailySchedule = ctrl(self,'m_ListDailySchedule')
                self.DSCursor = None
                self.ResultGrid = []
                self.HeaderSortReversed = False
                self.OrderCriteria = " ORDER BY RecordModificationDate DESC"

                self.Headers = ["No.","Record Key" ,"Project Type","Quote Number","Rev","Date Received","Date Expected",
                                                "Appl Engr","Salesperson", "Customer Key","Ship TO" ]

                self.ColumnCount = len(self.Headers)
                for i in range(0, self.ColumnCount):
                                self.m_ListDailySchedule.InsertColumn(i, self.Headers[i])
                                self.m_ListDailySchedule.SetColumnWidth(i, wx.LIST_AUTOSIZE_USEHEADER)
                self.m_ListDailySchedule.SetColumnWidth(0, 50)


                #bindings
                self.Bind(wx.EVT_CLOSE, self.on_close_frame)
                self.Bind(wx.EVT_TOOL, self.OnBtnSchedulePrint, id=xrc.XRCID('m_BtnSchedulePrint'))
                self.Bind(wx.EVT_TOOL, self.OnBtnScheduleExcel, id=xrc.XRCID('m_BtnScheduleExcel'))
                self.m_ListDailySchedule.Bind( wx.EVT_LIST_ITEM_ACTIVATED, self.OnClickScheduleItem)
                self.Bind(wx.EVT_LIST_COL_CLICK, self.OnHeaderSort,id=xrc.XRCID('m_ListDailySchedule') )


        def on_close_frame(self, event):
                self.Destroy()

        #takes in a python date and returns its string
        def GetDateString(self,d):
                try:
                        dt_str = str(d.date())
                        return dt_str
                except:
                        return ""

        def GetDailyRecords(self):

                # Get records with no completion date "AND" that are not more than 6 months overdue. More than 6 months overdue
                # assumed to be an abandoned job / record
                todayDate = datetime.datetime.today().date()
                SixMonth = todayDate - datetime.timedelta(days=180)
                SixMonthOldRequest = str(SixMonth)

                sql = "SELECT * FROM dbo.QuoteMaker WHERE (DateComp is NULL) OR (DateComp >= \'1/1/9999\') AND (DateRequest > \'"+ SixMonthOldRequest + "\') " + self.OrderCriteria
                self.DSCursor.execute(sql)
                row = self.DSCursor.fetchone()
                NumRes = 0
                SearchRecordKeys = []
                del self.ResultGrid[:]

                while row != None:

                                GridRow = [str(NumRes), str(row.RecordKey),str(row.ProjectType),str(row.QuoteNumber),str(row.RevLevel),
                                                   self.GetDateString(row.DateReceived),self.GetDateString(row.DateRequest),str(row.Assigned),
                                                   str(row.Saleperson),str(row.CustomerKey),str(row.ShipTO).strip()]

                                self.ResultGrid.append(GridRow)
                                SearchRecordKeys.append(row.RecordKey)

                                row = self.DSCursor.fetchone()
                                NumRes = NumRes + 1

                Title = "Daily Schedule: " + str(NumRes) + " open records"
                self.SetTitle(Title)

                if NumRes == 0: return   #No search results found

                #Get the parent class DBRecordKeys and replace all keys with just the daily schedule keys so we can browse through
                #the daily schedule records in the main window
                self.GetParent().DBRecordKeys = SearchRecordKeys
                self.GetParent().Index = 0
                self.GetParent().LoadRecord()

                #Insert results into list  control
                for x in range(0, NumRes):

                        self.m_ListDailySchedule.InsertStringItem(x,self.ResultGrid[x][0])

                        for col in range(1, self.ColumnCount):
                                self.m_ListDailySchedule.SetStringItem(x,col, self.ResultGrid[x][col])

                        self.RowColor(x)

                #Autosize the daily schedule list columns
                for col in range(0, self.ColumnCount):
                                self.m_ListDailySchedule.SetColumnWidth(col, wx.LIST_AUTOSIZE_USEHEADER)

        def RowColor(self, row):
                if row% 2:self.m_ListDailySchedule.SetItemBackgroundColour(row, "##fffff5")


   # Sort the results according to the column header clicked by the user
        def OnHeaderSort(self, event):
                column = event.GetColumn()
                self.ResultGrid.sort(key=lambda k: self.ConvColFormat(k[column]), reverse=self.HeaderSortReversed)
                self.HeaderSortReversed = not (self.HeaderSortReversed)
                self.m_ListDailySchedule.DeleteAllItems()
                SearchRecordKeys = []

           #Insert results into list  control
                NumRes = len(self.ResultGrid)  #should be equal to NumSelCustomers
                for x in range(0, NumRes):

                                self.m_ListDailySchedule.InsertStringItem(x,self.ResultGrid[x][0])

                                for col in range(1, self.ColumnCount):
                                                self.m_ListDailySchedule.SetStringItem(x,col, self.ResultGrid[x][col])

                                SearchRecordKeys.append(self.ResultGrid[x][1])


                self.GetParent().DBRecordKeys = SearchRecordKeys
                self.GetParent().Index = 0
                self.GetParent().LoadRecord()


        # try converting column format from string to float (for list sorting purposes)
        # if not succcessful (such as for customer name column), then return the original string back
        def ConvColFormat(self, str_val):
                try:
                        s = str_val
                        if s == None or len(s) == 0: return 0.0
                        s = s.replace('$','')
                        s = s.replace('%','')
                        return float(s)

                except:
                        return str_val



        #takes the daily schedule records and dumps output into excel
        def OnBtnScheduleExcel(self, event):
                excel = Dispatch('Excel.Application')
                excel.Visible = True
                wb = excel.Workbooks.Add()

                wb.ActiveSheet.Cells(1, 1).Value = 'Transferring data to Excel...'
                wb.ActiveSheet.Columns(1).AutoFit()

                R = len(self.ResultGrid)
                C = self.ColumnCount

                #Write the header
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(1,C))
                excel_range.Value = self.Headers

                #specify the excel range that the main data covers
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(2, 1),wb.ActiveSheet.Cells(R+1,C))
                excel_range.Value = self.ResultGrid

                #Autofit the columns
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(R+1,C))
                excel_range.Columns.AutoFit()

        def OnBtnSchedulePrint(self, event):

                html = '''<table border="1" cellspacing="0">\n'''
                R = len(self.ResultGrid)
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
                                html += '''<td align="left" valign="top" nowrap>{}&nbsp;</td>\n'''.format(self.ResultGrid[row][col])
                        html += '</tr>'

                html += '</table>'

                printer = HtmlEasyPrinting()

                #override these settings if you want
                printer_header = 'Printed on @DATE@, Page @PAGENUM@ of @PAGESCNT@'
                printer_font_size = 8
                printer_paper_type = wx.PAPER_LETTER        #wx.PAPER_11X17
                printer_paper_orientation = wx.LANDSCAPE        #wx.PORTRAIT
                printer_paper_margins = (1, 1, 1, 1)        #specify like (0, 0, 0, 0), otherwise None means whatever default is

                printer.SetHeader(printer_header)
                printer.SetStandardFonts(printer_font_size)
                printer.GetPrintData().SetPaperId(printer_paper_type)
                printer.GetPrintData().SetOrientation(printer_paper_orientation)

                if printer_paper_margins:
                        printer.GetPageSetupData().SetMarginTopLeft((printer_paper_margins[0], printer_paper_margins[1]))
                        printer.GetPageSetupData().SetMarginBottomRight((printer_paper_margins[2], printer_paper_margins[3]))

                printer.PrintText(html)

        def OnClickScheduleItem(self, event):
                item = self.m_ListDailySchedule.GetNextSelected(-1)
                RecordKey = self.m_ListDailySchedule.GetItemText(item, 1)

                #find the record key in the list
                i = self.GetParent().DBRecordKeys.index(RecordKey)

                #jump to this record on the main window
                if i < 0:
                                return
                else:
                                self.GetParent().Index = i
                                self.GetParent().LoadRecord()
                                self.Destroy()


