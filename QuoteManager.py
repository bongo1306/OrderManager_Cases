
import wx
from wx import xrc
ctrl = xrc.XRCCTRL

import Database as db
import pyodbc as mypyodbc
from QuoteSearch import QuoteSearchDialog
from QuoteSearch import DailyScheduleFrame
from Metrics import MetricsDialog
from UserManager import UserManagerDialog
from wx.html import HtmlEasyPrinting
import General as gn
import workdays
import datetime
import os
import sys
import threading
import subprocess
from CustomMessage import MessageDialog
import time
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
import sys
import wx
DropShipOrderNumberold = ''
datejobcompletedold = ''
projecttypeold = ''
salesorderold = ''

class QuoteManagerTab(object):
    def init_QuoteManager_tab(self):

        reload(sys)
        sys.setdefaultencoding('Cp1252')

        # Connect Events
        self.Bind( wx.EVT_TOOL, self.OnPageSetup, id=xrc.XRCID('m_toolPageSetup'))
        self.Bind( wx.EVT_TOOL, self.OnPrintPreview, id=xrc.XRCID('m_toolPrintPreview'))
        self.Bind( wx.EVT_TOOL, self.OnPrintRecord, id=xrc.XRCID('m_toolPrint'))
        self.Bind( wx.EVT_TOOL, self.OnGetQuoteNumber, id=xrc.XRCID('m_toolQuoteNumber'))
        self.Bind( wx.EVT_TOOL, self.OnAddRecord, id=xrc.XRCID('m_toolRecordAdd'))
        self.Bind( wx.EVT_TOOL, self.OnDeleteRecord, id=xrc.XRCID('m_toolRecordDelete'))
        self.Bind( wx.EVT_TOOL, self.OnDuplicateRecord, id=xrc.XRCID('m_toolRecordDuplicate'))
        self.Bind( wx.EVT_TOOL, self.OnSaveRecord, id=xrc.XRCID('m_toolSave'))
        self.Bind( wx.EVT_TOOL, self.OnQuoteSearch, id=xrc.XRCID('m_toolSearch'))
        self.Bind( wx.EVT_TOOL, self.OnSearchClear, id=xrc.XRCID('m_toolSearchClear'))
        self.Bind( wx.EVT_TOOL, self.OnDailySchedule, id=xrc.XRCID('m_toolSchedule'))
        self.Bind( wx.EVT_TOOL, self.OnMetrics, id=xrc.XRCID('m_toolMetrics'))
        self.Bind( wx.EVT_TOOL, self.OnPreviousRecord, id=xrc.XRCID('m_toolBack'))
        self.Bind( wx.EVT_TOOL, self.OnNextRecord, id=xrc.XRCID('m_toolNext'))
        self.Bind( wx.EVT_TOOL, self.OnManageUsers, id=xrc.XRCID('m_toolUsers'))
        self.Bind( wx.EVT_TOOL, self.OnAboutBox, id=xrc.XRCID('m_toolAbout'))
        self.Bind( wx.EVT_TOOL, self.OnExitApp, id=xrc.XRCID('m_toolQuit'))

        self.Bind(wx.EVT_BUTTON, self.OnQuickQuoteSearch, id=xrc.XRCID('m_BtnFindQuote'))
        self.Bind(wx.EVT_TEXT_ENTER, self.OnQuickQuoteSearch, id=xrc.XRCID('m_TextQuoteNumber'))
        self.Bind(wx.EVT_BUTTON, self.OnQuickSOSearch, id=xrc.XRCID('m_BtnFindSO'))
        self.Bind(wx.EVT_TEXT_ENTER, self.OnQuickSOSearch, id=xrc.XRCID('m_TextSalesOrderNum'))
        self.Bind(wx.EVT_BUTTON, self.OnQuickAddressSearch, id=xrc.XRCID('m_BtnFindAddress'))
        self.Bind(wx.EVT_TEXT_ENTER, self.OnQuickAddressSearch, id=xrc.XRCID('m_TextShipAddress'))
        self.Bind(wx.EVT_BUTTON, self.OnQuickCMATSearch, id=xrc.XRCID('m_BtnFindCMAT'))
        self.Bind(wx.EVT_TEXT_ENTER, self.OnSeekRecord, id=xrc.XRCID('m_TextRecordNum'))
        self.Bind(wx.EVT_DIRPICKER_CHANGED, self.OnProjectDirectory, id=xrc.XRCID('m_BrowseDir'))
        self.Bind(wx.EVT_TEXT, self.OnComboCustomerSelected, id=xrc.XRCID('m_ComboCustomerName'))

        self.Bind(wx.EVT_CHECKBOX, self.OnCheckRefrigerant, id=xrc.XRCID('m_CheckCO2'))
        self.Bind(wx.EVT_CHECKBOX, self.OnCheckRefrigerant, id=xrc.XRCID('m_CheckAmmonia'))
        self.Bind(wx.EVT_CHECKBOX, self.OnCheckRefrigerant, id=xrc.XRCID('m_CheckGlycol'))


        #self.m_checkListBox = wx.FindWindowByName('m_checkListBox')
        self.m_ComboProjectType = wx.FindWindowByName('m_ComboProjectType')
        self.m_ComboBidOpen = wx.FindWindowByName('m_ComboBidOpen')
        self.m_ComboProjectStatus = wx.FindWindowByName('m_ComboProjectStatus')
        self.m_ComboZone = wx.FindWindowByName('m_ComboZone')
        self.m_TextQuoteNumber = wx.FindWindowByName('m_TextQuoteNumber')
        self.m_BtnFindQuote = wx.FindWindowByName('m_BtnFindQuote')
        self.m_ComboRevLevel = wx.FindWindowByName('m_ComboRevLevel')
        self.m_TextSalesOrderNum = wx.FindWindowByName('m_TextSalesOrderNum')
        self.m_BtnFindSO = wx.FindWindowByName('m_BtnFindSO')
        self.m_TextDropShipOrderNum = wx.FindWindowByName('m_TextDropShipOrderNum')
        self.m_DateReceived = wx.FindWindowByName('m_DateReceived')
        self.m_DateDue = wx.FindWindowByName('m_DateDue')
        self.m_DateCompleted = wx.FindWindowByName('m_DateCompleted')
        self.m_TextTurnAround = wx.FindWindowByName('m_TextTurnAround')
        self.m_ComboAssigned = wx.FindWindowByName('m_ComboAssigned')
        self.m_ComboSP = wx.FindWindowByName('m_ComboSP')
        self.m_TextEquipUSD = wx.FindWindowByName('m_TextEquipUSD')
        self.m_TextBODollars = wx.FindWindowByName('m_TextBODollars')
        self.m_TextBOPercent = wx.FindWindowByName('m_TextBOPercent')
        self.m_TextMultiplier = wx.FindWindowByName('m_TextMultiplier')
        self.m_BrowseDir =  wx.FindWindowByName('m_BrowseDir')
        self. m_TextProjectFolder = wx.FindWindowByName('m_TextProjectFolder')
        self.m_ComboCustomerName = wx.FindWindowByName('m_ComboCustomerName')
        self.m_TextCustomerKey  = wx.FindWindowByName('m_TextCustomerKey')
        self.m_TextCustomerNumber = wx.FindWindowByName('m_TextCustomerNumber')
        self.m_TextShipAddress = wx.FindWindowByName('m_TextShipAddress')
        self.m_TextNotes = wx.FindWindowByName('m_TextNotes')
        self.m_CheckCO2 = wx.FindWindowByName('m_CheckCO2')
        self.m_CheckAmmonia = wx.FindWindowByName('m_CheckAmmonia')
        self.m_CheckGlycol = wx.FindWindowByName('m_CheckGlycol')
        self.m_ComboCMAT = wx.FindWindowByName('m_ComboCMAT')
        #self.m_TextCMAT = wx.FindWindowByName('m_TextCMAT')
        self.m_BtnFindCMAT = wx.FindWindowByName('m_BtnFindCMAT')
        self.m_TextRecordInfo = wx.FindWindowByName('m_TextRecordInfo')
        self.m_TextRecordNum = wx.FindWindowByName('m_TextRecordNum')
        self.m_TextRecordCount = wx.FindWindowByName('m_TextRecordCount')
        self.mainWnd = wx.FindWindowByName('frame:main')

        #Database connection variables
        self.AENames = []
        self.SPNames = []
        self.CMATNames = []
        self.CustomerData = []
        self.CustomerNumbers = []
        self.DBRecordKeys = []
        ##self.EmailIDS = []
        self.count = 0
        self.Index = 0                              # zero based index of current record displayed
        self.dbCursor = None
        self.dbCursor_threaded = None
        threading.Thread(target=self.FetchDB).start()

        conn = db.connect_to_eng04_database() #create a new database connection for QuoteManager.
        self.dbCursor = conn.cursor()

        #self.mainWnd.Maximize()

        #bind various individual controls to theor events
        self.m_TextProjectFolder.SetCursor(wx.StockCursor(wx.CURSOR_CLOSED_HAND))
        self.m_TextProjectFolder.Bind(wx.EVT_LEFT_UP, self.OnClickProjectDir)


        # initilize the printing variables
        self.PrintText = ""
        self.printer_font_size = 10
        self.printer_paper_type = wx.PAPER_11X17          #wx.PAPER_LETTER
        self.printer_paper_orientation = wx.LANDSCAPE       #wx.LANDSCAPE
        self.margins = (15, 15, 15, 15)                     #specify like (0, 0, 0, 0), otherwise None means whatever default is


    def OnPreviousRecord(self, event):

        self.Index = self.Index - 1
        self.LoadRecord()

    def OnNextRecord(self,event):

        self.Index = self.Index +  1
        self.LoadRecord()

    def OnSeekRecord(self, event):
        #check if value is numeric
        indexString = self.m_TextRecordNum.GetValue()

        if indexString.isdigit() == False:
            return

        #convert to integer
        index = int(indexString)
        index = index - 1

        #check if index is within proper bounds
        if index < 0 or index >= self.count: return
        else:
            self.Index = index
            self.LoadRecord()


    # Fetch all records from the database. This function will be called from a thread
    def FetchDB(self):
        time.sleep(2)

        #Database connection variables
        conn = db.connect_to_eng04_database() #create a new database connection for QuoteManager.
        self.dbCursor_threaded = conn.cursor()

        self.FillAEComboBoxNames()    # Fill in the names of all application engineers in combo box
        self.FillSPComboBoxNames()    # Fill in the names of all saleperson in the combo box
        self.FillCustomerComboBox()    # Fill in the names of all customers in the combo box
        self.FillCMATComboBox()        # Fill in the names of all CMAT associated with alternate refrigerants
        self.FillDBRecordKeys()
        ##self.FillEMailCheckBoxList()

    def FillDBRecordKeys(self):

        del self.DBRecordKeys[:]
        self.DBRecordKeys = []
        self.dbCursor_threaded.execute('SELECT TOP 5000 RecordKey FROM dbo.QuoteMaker ORDER BY RecordModificationDate DESC, RecordModificationTime DESC')
        row = self.dbCursor_threaded.fetchone()

        while row != None:
                self.DBRecordKeys.append(row.RecordKey)
                row = self.dbCursor_threaded.fetchone()

        self.DBRecordKeys.reverse()
        self.count = len(self.DBRecordKeys)
        self.Index = self.count-1
        self.LoadRecord()


    # Load the record at the current index self.Index
    def LoadRecord(self):

        #Before loading record make sure it is within proper bounds
        self.count = len(self.DBRecordKeys)
        if self.count == 0: return  #No records to display

        if self.Index < 0: self.Index = 0
        elif self.Index > self.count - 1:  self.Index = self.count - 1

        key = self.DBRecordKeys[self.Index]
        sql = "SELECT * FROM dbo.QuoteMaker WHERE RecordKey=\'" + key + "\'"
        self.dbCursor.execute(sql)
        DBRecord = self.dbCursor.fetchone()

        #update the project type ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectType = str(DBRecord.ProjectType)

        global projecttypeold
        projecttypeold = ProjectType
        #print(projecttypeold)
        if self.m_ComboProjectType.FindString(ProjectType) == wx.NOT_FOUND:
            self.m_ComboProjectType.Append(ProjectType)
        self.m_ComboProjectType.SetStringSelection(ProjectType)

        #update the bid open ("Yes" or "No") ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        BidOpen  = str(DBRecord.BidOpen)
        BidOpen = BidOpen.strip()

        if len(BidOpen) < 1: BidOpen = "NA"

        if self.m_ComboBidOpen.FindString(BidOpen) == wx.NOT_FOUND:
            self.m_ComboBidOpen.Append(BidOpen)
        self.m_ComboBidOpen.SetStringSelection(BidOpen)

        # add the status of the project ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectStatus = str(DBRecord.ProjectStatus)
        ProjectStatus = ProjectStatus.strip()

        if len(ProjectStatus) < 1:
            ProjectStatus = "NA"

        if self.m_ComboProjectStatus.FindString(ProjectStatus) == wx.NOT_FOUND:
            self.m_ComboProjectStatus.Append(ProjectStatus)

        self.m_ComboProjectStatus.SetStringSelection(ProjectStatus)

        # add the zone of the project ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Zone= str(DBRecord.Zone)

        if self.m_ComboZone.FindString(Zone) == wx.NOT_FOUND:
            self.m_ComboZone.Append(Zone)
        self.m_ComboZone.SetStringSelection(Zone)

        # update the quote number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        QuoteNumber = str(DBRecord.QuoteNumber)
        self.m_TextQuoteNumber.SetLabelText(QuoteNumber)


        # update the quote revision level on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        RevLevel = str(DBRecord.RevLevel)

        if self.m_ComboRevLevel.FindString(RevLevel) == wx.NOT_FOUND:
            self.m_ComboRevLevel.Append(RevLevel)
        self.m_ComboRevLevel.SetStringSelection(RevLevel)


        # update the sales order number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        SalesOrderNum = str(DBRecord.SalesOrderNum)
        global salesorderold
        salesorderold = SalesOrderNum
        self.m_TextSalesOrderNum.SetLabelText(SalesOrderNum)

        # update the drop ship order number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~
        DropShipOrderNumber = str(DBRecord.DropShipOrderNum)
        global DropShipOrderNumberold
        DropShipOrderNumberold = DropShipOrderNumber
        self.m_TextDropShipOrderNum.SetLabelText(DropShipOrderNumber)

        # update date received ~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx_Rec = gn.PY2WX(DBRecord.DateReceived)
        self.m_DateReceived.SetValue(wx_Rec)

        # update date requested ~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx_Req = gn.PY2WX(DBRecord.DateRequest)
        self.m_DateDue.SetValue(wx_Req)

        # update date job completed ~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx_Comp = gn.PY2WX(DBRecord.DateComp)
        self.m_DateCompleted.SetValue(wx_Comp)

        global datejobcompletedold
        datejobcompletedold = str(wx_Comp)
        # calculate the turn around time in days ~~~~~~~~~~~~~~~~~~~~~~~~~~
        TADays = gn.GetTADays(DBRecord.DateReceived, DBRecord.DateComp)
        self.m_TextTurnAround.SetLabelText(TADays)

        # add the name of the appl engineer to the combo box and select it ~~~~~~~~~~~
        Assigned = str(DBRecord.Assigned)

        if self.m_ComboAssigned.FindString(Assigned) == wx.NOT_FOUND:
            self.m_ComboAssigned.Append(Assigned)
        self.m_ComboAssigned.SetStringSelection(Assigned)

        # add the name of the salesperson to the combo box and select it ~~~~~~~~~~~
        Salesperson = str(DBRecord.Saleperson)

        if self.m_ComboSP.FindString(Salesperson) == wx.NOT_FOUND:
            self.m_ComboSP.Append(Salesperson)
        self.m_ComboSP.SetStringSelection(Salesperson)

        # add the Equipment and Buyout to the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        EquipPrice = "0"
        if DBRecord.EquipPrice != None:
              EquipPrice = '{0:,}'.format(DBRecord.EquipPrice)
        self.m_TextEquipUSD.SetLabelText(EquipPrice)

        BuyoutPrice = "0"
        if DBRecord.BuyoutPrice != None:
              BuyoutPrice = '{0:,}'.format(DBRecord.BuyoutPrice)
        self.m_TextBODollars.SetLabelText(BuyoutPrice)

        BuyoutPercent = "0"
        if DBRecord.BuyoutPercent != None:
              BuyoutPercent = str("%0.0f" %DBRecord.BuyoutPercent)
        self.m_TextBOPercent.SetLabelText(BuyoutPercent)

        Multiplier = "0"
        if DBRecord.Multiplier != None:
              Multiplier = str("%0.2f" %DBRecord.Multiplier)
        self.m_TextMultiplier.SetLabelText(Multiplier)

        # Update the project folder
        ProjectFolder = ""
        if DBRecord.ProjectFolder != None:
              ProjectFolder = str(DBRecord.ProjectFolder)
        self.m_TextProjectFolder.SetLabelText(ProjectFolder)

        # add the customer name, number, key and ship to address to the display ~~~~~~~
        Customer = str(DBRecord.Customer)
        Customer = Customer.replace("&","and")

        if self.m_ComboCustomerName.FindString(Customer) == wx.NOT_FOUND:
            self.m_ComboCustomerName.Append(Customer)
        self.m_ComboCustomerName.SetStringSelection(Customer)

        CustomerKey = DBRecord.CustomerKey
        self.m_TextCustomerKey.SetLabelText(CustomerKey)

        CustomerNumber = str(DBRecord.CustomerNumber)
        self.m_TextCustomerNumber.SetLabelText(CustomerNumber)

        ShipTO = str(DBRecord.ShipTO)
        ShipTO = ShipTO.replace("&","and")
        self.m_TextShipAddress.SetLabelText(ShipTO)

        #Update the record information (box in lower right hand corner of the screen) ~~~~~
        # Update the record key
        RecordKey = str(DBRecord.RecordKey)

        # Record Creation Log
        RecordCreatorName = str(DBRecord.RecordCreatorName)
        RecordCreationDate = self.PY2STR_DATE(DBRecord.RecordCreationDate)
        RecordCreationTime = self.PY2STR_TIME(DBRecord.RecordCreationTime)

        #Record Modification Log
        RecordModifierName = str(DBRecord.RecordModifierName)
        RecordModificationDate = self.PY2STR_DATE(DBRecord.RecordModificationDate)
        RecordModificationTime = self.PY2STR_TIME(DBRecord.RecordModificationTime)

        #update notes and record information
        Comments = str(DBRecord.Comments)
        Comments = Comments.replace("&","and")

        #DrawingComments = str(DBRecord.DrawingComments)
        #if DBRecord.DrawingComments != None and len(DrawingComments) > 1:
        #    Comments = Comments + "\r\n\r\n\r\n" + DrawingComments

        self.m_TextNotes.SetLabelText(Comments)
        idx = self.Index + 1
        self.m_TextRecordNum.SetLabelText(str("%0.0f" %idx))
        self.m_TextRecordCount.SetLabelText(" of " + str(self.count))

        txt = "Record Key: " + RecordKey + "\r\n"
        #txt = txt + "Record Number: " + str("%0.0f" %i) + " of " + str(self.count) + "\r\n"
        txt = txt + "Record Creator Name: "+ RecordCreatorName + "\r\n"
        txt = txt + "Record Creation Date: " + RecordCreationDate + "\r\n"
        txt = txt + "Record Creation Time: " + RecordCreationTime + "\r\n"
        txt = txt + "Record Modifier Name: " + RecordModifierName + "\r\n"
        txt = txt + "Record Modification Date: " + RecordModificationDate + "\r\n"
        txt = txt + "Record Modification Time: " + RecordModificationTime + "\r\n"

        self.m_TextRecordInfo.SetLabelText (txt)

        #If any of the alternate refrigerants are checked, then make the CMAT drop down visible
        CarbonDioxide = DBRecord.CarbonDioxide
        self.m_CheckCO2.SetValue(False)
        if CarbonDioxide != None: self.m_CheckCO2.SetValue(CarbonDioxide)

        Ammonia = DBRecord.Ammonia
        self.m_CheckAmmonia.SetValue(False)
        if Ammonia != None: self.m_CheckAmmonia.SetValue(Ammonia)

        Glycol = DBRecord.Glycol
        self.m_CheckGlycol.SetValue(False)
        if Glycol != None: self.m_CheckGlycol.SetValue(Glycol)

        CMAT = DBRecord.CMAT
        if CMAT != None and len(CMAT) > 1:
            if self.m_ComboCMAT.FindString(CMAT) == wx.NOT_FOUND: self.m_ComboCMAT.Append(CMAT)
            self.m_ComboCMAT.SetStringSelection(CMAT)
        else:
            self.m_ComboCMAT.SetStringSelection("ANY")

        if CarbonDioxide != True and Ammonia != True and Glycol != True: self.HideCMAT()
        else: self.ShowCMAT()


    def ShowCMAT(self):
            self.m_ComboCMAT.Show()
            self.m_BtnFindCMAT.Show()
            #self.m_TextCMAT.Show()
            self.mainWnd.Maximize()
            self.mainWnd.Restore()


    def HideCMAT(self):
            self.m_ComboCMAT.Hide()
            self.m_BtnFindCMAT.Hide()
            #self.m_TextCMAT.Hide()


    def OnProjectDirectory(self, event):
        self.ProjectFolder = self.m_BrowseDir.GetPath()
        self.m_TextProjectFolder.SetLabelText(self.ProjectFolder)

    def OnClickProjectDir(self, event):
        dir = self.m_TextProjectFolder.GetValue()
        try:
            os.startfile(dir)
        except:
            wx.MessageBox("Unable to open project folder", "Error Opening Poject Folder",wx.OK | wx.ICON_EXCLAMATION)

    def FillAEComboBoxNames(self):
        time.sleep(2)
        del self.AENames[:]
        self.dbCursor_threaded.execute('SELECT * FROM dbo.ApplicationsTable')
        row = self.dbCursor_threaded.fetchone()

        while row != None:
            self.AENames.append(row.Name)
            row = self.dbCursor_threaded.fetchone()

        self.m_ComboAssigned.SetItems(self.AENames)

    def FillEMailCheckBoxList(self):
        time.sleep(1)
        del self.EmailIDS[:]
        self.dbCursor_threaded.execute("SELECT * FROM dbo.emailforquote where etype = 'optional' and enabled = 1")
        row = self.dbCursor_threaded.fetchone()

        while row != None:
            self.EmailIDS.append(row.name)
            row = self.dbCursor_threaded.fetchone()
        self.m_checkListBox.Set(self.EmailIDS)

    def FillSPComboBoxNames(self):
        time.sleep(1)
        del self.SPNames[:]
        self.dbCursor_threaded.execute('SELECT * FROM dbo.SalepersonTable')
        row = self.dbCursor_threaded.fetchone()

        while row != None:
            self.SPNames.append(row.Name)
            row = self.dbCursor_threaded.fetchone()

        self.m_ComboSP.SetItems(self.SPNames)



    def FillCustomerComboBox(self):

        del self.CustomerData[:]
        del self.CustomerNumbers[:]
        self.dbCursor_threaded.execute('SELECT * FROM dbo.CustomerTable')
        row = self.dbCursor_threaded.fetchone()

        while row != None:
            self.CustomerData.append(row.Name)
            self.CustomerNumbers.append(row.Number)
            row = self.dbCursor_threaded.fetchone()

        self.m_ComboCustomerName.SetItems(self.CustomerData)

    def FillCMATComboBox(self):

        del self.CMATNames[:]
        self.CMATNames.append("ANY")
        self.dbCursor_threaded.execute('SELECT * FROM dbo.AlternateCMAT')
        row = self.dbCursor_threaded.fetchone()

        while row != None:
            self.CMATNames.append(row.Name)
            row = self.dbCursor_threaded.fetchone()

        self.m_ComboCMAT.SetItems(self.CMATNames)


    def OnQuoteSearch( self, event ):
        searchWnd = QuoteSearchDialog(self)
        searchWnd.SearchCursor = self.dbCursor
        searchWnd.PopulateAE(self.AENames)
        searchWnd.PopulateSP(self.SPNames)
        searchWnd.PopulateCMAT(self.CMATNames)
        searchWnd.OnCheckSearchCMAT(None)
        searchWnd.SetSize((1320, 780))
        #searchWnd.Maximize()
        searchWnd.ShowModal()


    #Search quote number directly from the main window
    def OnQuickQuoteSearch(self, event ):

        QuoteNumber = str(self.m_TextQuoteNumber.GetValue())

        if QuoteNumber.find("update") != -1:
            threading.Thread(target=self.OnNightlyBuild).start()
            return

        if len(QuoteNumber) < 1 or QuoteNumber.isdigit() == False:
             wx.MessageBox("Please select a valid QuoteNumber", "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        sql = "SELECT * FROM dbo.QuoteMaker WHERE QuoteNumber=" + QuoteNumber
        self.dbCursor.execute(sql)
        SearchRecordKeys = []
        row = self.dbCursor.fetchone()

        while row != None:
                                SearchRecordKeys.append(row.RecordKey)
                                row = self.dbCursor.fetchone()

        NumRes = len(SearchRecordKeys)
        if NumRes == 0:   #No searh results found
                msg = "No results found for " + QuoteNumber
                wx.MessageBox(msg, "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                return
        else:
            self.DBRecordKeys = SearchRecordKeys
            self.Index = 0
            self.LoadRecord()

    #Search ship address directly from the main window
    def OnQuickAddressSearch(self, event ):

        ShipAddress = str(self.m_TextShipAddress.GetValue())

        if len(ShipAddress) < 2:
             wx.MessageBox("Please provide a valid ship address or store number for search",
                           "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        Address = "%" + ShipAddress + "%"
        sql = "SELECT * FROM dbo.QuoteMaker WHERE (ShipTO LIKE ?) ORDER BY RecordModificationDate DESC, RecordModificationTime DESC"
        self.dbCursor.execute(sql, Address)
        SearchRecordKeys = []
        row = self.dbCursor.fetchone()

        while row != None:
                                SearchRecordKeys.append(row.RecordKey)
                                row = self.dbCursor.fetchone()

        NumRes = len(SearchRecordKeys)
        if NumRes == 0:   #No searh results found
                msg = "No results found for " + ShipAddress
                wx.MessageBox(msg, "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                return
        else:
            self.DBRecordKeys = SearchRecordKeys
            self.Index = 0
            self.LoadRecord()

    #Search sales order number directly from the main window
    def OnQuickSOSearch(self, event ):

        SO = str(self.m_TextSalesOrderNum.GetValue())
        SO = SO.strip()
        if len(SO) <= 1:
             wx.MessageBox("Please select a valid Sales Order Number", "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        sql = "SELECT * FROM dbo.QuoteMaker WHERE SalesOrderNum LIKE \'%" + SO + "%\'"
        self.dbCursor.execute(sql)
        SearchRecordKeys = []
        row = self.dbCursor.fetchone()

        while row != None:
                                SearchRecordKeys.append(row.RecordKey)
                                row = self.dbCursor.fetchone()

        NumRes = len(SearchRecordKeys)
        if NumRes == 0:   #No searh results found
                msg = "No results found for " + SO
                wx.MessageBox(msg, "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                return
        else:
            self.DBRecordKeys = SearchRecordKeys
            self.Index = 0
            self.LoadRecord()


    #Search alternate refrigerants CMATs directly from the main window
    def OnQuickCMATSearch(self, event ):

        CM = str(self.m_ComboCMAT.GetValue())
        CM = CM.strip()
        if len(CM) <= 1:
             wx.MessageBox("Please select a valid CMAT", "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        if CM.find("ANY") == -1:
            sql = "SELECT * FROM dbo.QuoteMaker WHERE CMAT=\'" + CM + "\'"
        else:
            sql = "SELECT * FROM dbo.QuoteMaker WHERE CMAT IS NOT NULL"
        self.dbCursor.execute(sql)
        SearchRecordKeys = []
        row = self.dbCursor.fetchone()

        while row != None:
                                SearchRecordKeys.append(row.RecordKey)
                                row = self.dbCursor.fetchone()

        NumRes = len(SearchRecordKeys)
        if NumRes == 0:   #No searh results found
                msg = "No results found for " + CM
                wx.MessageBox(msg, "Search Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                return
        else:
            self.DBRecordKeys = SearchRecordKeys
            self.Index = 0
            self.LoadRecord()


    def OnSearchClear( self, event ):
        self.FillDBRecordKeys()


    def OnDailySchedule( self, event ):
        ds_frame = DailyScheduleFrame(self)
        ds_frame.DSCursor = self.dbCursor
        ds_frame.GetDailyRecords()
        ds_frame.Show()
        ds_frame.Maximize()

    def OnMetrics( self, event ):
        metricsWnd = MetricsDialog(self)
        metricsWnd.SearchCursor = self.dbCursor
        metricsWnd.PopulateLists()
        metricsWnd.Maximize()
        metricsWnd.ShowModal()


    def OnPageSetup(self, event):

        dlg = wx.PageSetupDialog(self)
        S = dlg.PageSetupDialogData

        S.PrintData.PaperId = self.printer_paper_type
        S.PrintData.SetOrientation(self.printer_paper_orientation)
        S.SetMarginTopLeft((self.margins[0], self.margins[1]))
        S.SetMarginBottomRight((self.margins[2], self.margins[3]))

        if dlg.ShowModal() == wx.ID_OK:

                self.printer_paper_type = S.PaperId
                self.printer_paper_orientation = S.PrintData.GetOrientation()

                self.margins = (S.GetMarginTopLeft().x,S.GetMarginTopLeft().y,
                                S.GetMarginBottomRight().x,S.GetMarginBottomRight().y)


    def OnPrintRecord(self, evt):

        printer = HtmlEasyPrinting()
        printer.SetParentWindow(self)

        printer.SetHeader('Printed on @DATE@, Page @PAGENUM@ of @PAGESCNT@')
        printer.SetStandardFonts(self.printer_font_size)
        printer.GetPrintData().SetPaperId(self.printer_paper_type)
        printer.GetPrintData().SetOrientation(self.printer_paper_orientation)

        printer.GetPageSetupData().SetMarginTopLeft((self.margins[0], self.margins[1]))
        printer.GetPageSetupData().SetMarginBottomRight((self.margins[2], self.margins[3]))

        self.CreateScreenShotFile()
        printer.PrintFile (gn.resource_path('screenshot.htm'))

    def OnPrintPreview(self, event):

        printer = HtmlEasyPrinting()
        printer.SetParentWindow(self)

        printer.SetHeader('Printed on @DATE@, Page @PAGENUM@ of @PAGESCNT@')
        printer.SetStandardFonts(self.printer_font_size)
        printer.GetPrintData().SetPaperId(self.printer_paper_type)
        printer.GetPrintData().SetOrientation(self.printer_paper_orientation)

        printer.GetPageSetupData().SetMarginTopLeft((self.margins[0], self.margins[1]))
        printer.GetPageSetupData().SetMarginBottomRight((self.margins[2], self.margins[3]))

        self.CreateScreenShotFile()
        printer.PreviewFile(gn.resource_path('screenshot.htm'))

    #Creates a html file "screenshot.htm" that contains the screenshot png image
    def CreateScreenShotFile(self):

        rect = self.mainWnd.GetRect()
        dcScreen = wx.ScreenDC()

        bmp = wx.EmptyBitmap(rect.width, rect.height)

        memDC = wx.MemoryDC()
        memDC.SelectObject(bmp)
        memDC.Blit( 0, 0, rect.width, rect.height,dcScreen, rect.x,  rect.y)
        memDC.SelectObject(wx.NullBitmap)

        img = bmp.ConvertToImage()
        fileName = "screenshot.png"
        img.SaveFile(gn.resource_path(fileName), wx.BITMAP_TYPE_PNG)

        html = "<html>\n<body>\n<center><img src=screenshot.png></center>\n</body>\n</html>"
        f = file(gn.resource_path('screenshot.htm'), 'w')
        f.write(html)
        f.close()


    def OnGetQuoteNumber(self, event):
        #Get the next available quote number ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        sql = 'SELECT MAX(QuoteNumber) AS QMax FROM dbo.QuoteMaker'
        self.dbCursor.execute(sql)
        row = self.dbCursor.fetchone()

        if row == None: return

        NextQuoteNumber = row.QMax + 1
        msg = "Please use the next available Quote # " + str(NextQuoteNumber)
        title = "New Quote"
        wx.MessageBox(msg, title, wx.OK|wx.ICON_INFORMATION)


    def OnAddRecord(self, event):

        #if read only mode then exit
        if gn.mode_readonly == 'True':  #read-only mode is true, which means no deletion permitted
            wx.MessageBox("You cannot make changes to the program while in read-only mode","Add Unsuccessful",wx.OK | wx.ICON_INFORMATION)
            return

        #update the project type ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectType = ""

        if self.m_ComboProjectType.FindString(ProjectType) == wx.NOT_FOUND:
            self.m_ComboProjectType.Append(ProjectType)
        self.m_ComboProjectType.SetStringSelection(ProjectType)

        #update the bid open ("Yes" or "No") ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        BidOpen  = "NA"

        if self.m_ComboBidOpen.FindString(BidOpen) == wx.NOT_FOUND:
            self.m_ComboBidOpen.Append(BidOpen)
        self.m_ComboBidOpen.SetStringSelection(BidOpen)

        # add the status of the project ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectStatus = "NA"

        if self.m_ComboProjectStatus.FindString(ProjectStatus) == wx.NOT_FOUND:
            self.m_ComboProjectStatus.Append(ProjectStatus)
        self.m_ComboProjectStatus.SetStringSelection(ProjectStatus)

        # add the zone of the project ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Zone= ""

        if self.m_ComboZone.FindString(Zone) == wx.NOT_FOUND:
            self.m_ComboZone.Append(Zone)
        self.m_ComboZone.SetStringSelection(Zone)

        # update the quote number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        QuoteNumber = ""
        self.m_TextQuoteNumber.SetLabelText(QuoteNumber)


        # update the quote revision level on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        RevLevel = ""

        if self.m_ComboRevLevel.FindString(RevLevel) == wx.NOT_FOUND:
            self.m_ComboRevLevel.Append(RevLevel)
        self.m_ComboRevLevel.SetStringSelection(RevLevel)


        # update the sales order number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        SalesOrderNum = ""
        self.m_TextSalesOrderNum.SetLabelText(SalesOrderNum)

        # update the drop ship order number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~
        DropShipOrderNumber = ""
        self.m_TextDropShipOrderNum.SetLabelText(DropShipOrderNumber)

        # update date received ~~~~~~~~~~~~~~~~~~~~~~~~~~
        self.m_DateReceived.SetValue(wx.DateTime_Today())
        self.m_DateReceived.SetValue(wx.DefaultDateTime)

        # update date requested ~~~~~~~~~~~~~~~~~~~~~~~~~~
        self.m_DateDue.SetValue(wx.DateTime_Today())
        self.m_DateDue.SetValue(wx.DefaultDateTime)

        # update date job completed ~~~~~~~~~~~~~~~~~~~~~~~~~~
        self.m_DateCompleted.SetValue(wx.DateTime_Today())
        self.m_DateCompleted.SetValue(wx.DefaultDateTime)

        # calculate the turn around time in days ~~~~~~~~~~~~~~~~~~~~~~~~~~
        self.m_TextTurnAround.SetLabelText("")

        # add the name of the appl engineer to the combo box and select it ~~~~~~~~~~~
        Assigned = ""

        if self.m_ComboAssigned.FindString(Assigned) == wx.NOT_FOUND:
            self.m_ComboAssigned.Append(Assigned)
        self.m_ComboAssigned.SetStringSelection(Assigned)

        # add the name of the salesperson to the combo box and select it ~~~~~~~~~~~
        Salesperson = ""

        if self.m_ComboSP.FindString(Salesperson) == wx.NOT_FOUND:
            self.m_ComboSP.Append(Salesperson)
        self.m_ComboSP.SetStringSelection(Salesperson)

        # add the Equipment and Buyout to the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        EquipPrice = ""
        self.m_TextEquipUSD.SetLabelText(EquipPrice)

        BuyoutPrice = ""
        self.m_TextBODollars.SetLabelText(BuyoutPrice)

        BuyoutPercent = ""
        self.m_TextBOPercent.SetLabelText(BuyoutPercent)

        Multiplier = ""
        self.m_TextMultiplier.SetLabelText(Multiplier)

        # Update the project folder
        ProjectFolder = ""
        self.m_TextProjectFolder.SetLabelText(ProjectFolder)

        # add the customer name, number, key and ship to address to the display ~~~~~~~
        Customer = ""
        self.m_ComboCustomerName.SetLabelText(Customer)

        CustomerKey = ""
        self.m_TextCustomerKey.SetLabelText(CustomerKey)

        CustomerNumber = ""
        self.m_TextCustomerNumber.SetLabelText(CustomerNumber)

        ShipTO = ""
        self.m_TextShipAddress.SetLabelText(ShipTO)

        #Update the record information (box in lower right hand corner of the screen) ~~~~~
        # Update the record key
        RecordKey = ""

        # Record Creation Log
        RecordCreatorName = ""
        RecordCreationDate = ""
        RecordCreationTime = ""

        #Record Modification Log
        RecordModifierName = ""
        RecordModificationDate = ""
        RecordModificationTime =  ""

        #update notes and record information
        Comments = ""
        self.m_TextNotes.SetLabelText(Comments)

        self.m_TextRecordNum.SetLabelText("")
        self.m_TextRecordCount.SetLabelText(" of " + str(self.count))

        txt = "Record Key: " + RecordKey + "\r\n"
        txt = txt + "Record Creator Name: "+ RecordCreatorName + "\r\n"
        txt = txt + "Record Creation Date: " + RecordCreationDate + "\r\n"
        txt = txt + "Record Creation Time: " + RecordCreationTime + "\r\n"
        txt = txt + "Record Modifier Name: " + RecordModifierName + "\r\n"
        txt = txt + "Record Modification Date: " + RecordModificationDate + "\r\n"
        txt = txt + "Record Modification Time: " + RecordModificationTime + "\r\n"

        self.m_TextRecordInfo.SetLabelText (txt)

    def OnDeleteRecord(self, event):
        #if read only mode then exit
        if gn.mode_readonly == 'True':  #read-only mode is true, which means no deletion permitted
            wx.MessageBox("You cannot make changes to the program while in read-only mode","Delete Unsuccessful",wx.OK | wx.ICON_INFORMATION)
            return

        #get the project type ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectType = self.m_ComboProjectType.GetValue()

        # get the quote number ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        QuoteNumber = str(self.m_TextQuoteNumber.GetValue())

        # get the quote revision level ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        RevLevel = str(self.m_ComboRevLevel.GetValue())

        # get the record key
        RecordKey = QuoteNumber + ProjectType + RevLevel

        #Check to see if the record key exists in the database
        sql = "SELECT * FROM dbo.QuoteMaker WHERE RecordKey=\'" + RecordKey + "\'"
        self.dbCursor.execute(sql)
        row = self.dbCursor.fetchone()
        if row == None: return

        #confirm delete
        msg = ProjectType + " " + QuoteNumber  + " Revision " + RevLevel + " wil be permanently deleted. Continue with deletion?"
        ContinueDelete = wx.MessageBox(msg, "Delete Confirmation",wx.YES_NO | wx.ICON_INFORMATION)

        if ContinueDelete == wx.NO:
            return

        # Get the record creator name and the name of the logged in user
        curUser = str(gn.user).strip()
        recordCreator = str(row.RecordCreatorName).strip()

        #Check the delete pass phrase. Either passphrase has to be correct or the user deleting the record has to be the
        #original record creator
        notesText = self.m_TextNotes.GetValue()
        if notesText.find("superadmin") == -1 and curUser.find(recordCreator) == -1:
            MsgWnd = MessageDialog(self)
            MsgWnd.ShowModal()
            return

        #delete the record database
        sql = "DELETE FROM dbo.QuoteMaker WHERE RecordKey=\'" + RecordKey + "\'"
        self.dbCursor.execute(sql)
        self.dbCursor.commit()

        #delete the record from the DBRecordKeys
        self.DBRecordKeys.remove(RecordKey)
        self.LoadRecord()

        wx.MessageBox("The record was successfully deleted","Delete Successful",wx.OK | wx.ICON_INFORMATION)


    def OnDuplicateRecord(self, event):

        #if read only mode then exit
        if gn.mode_readonly == 'True':  #read-only mode is true, which means no deletion permitted
            wx.MessageBox("You cannot make changes to the program while in read-only mode","Duplication Unsuccessful",wx.OK | wx.ICON_INFORMATION)
            return

        # update the quote number on the display ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        QuoteNumber = self.m_TextQuoteNumber.GetLabelText() + "-COPY"
        self.m_TextQuoteNumber.SetLabelText(QuoteNumber)

        #Update the record information (box in lower right hand corner of the screen) ~~~~~
        # Update the record key
        RecordKey = ""

        # Record Creation Log
        RecordCreatorName = ""
        RecordCreationDate = ""
        RecordCreationTime = ""

        #Record Modification Log
        RecordModifierName = ""
        RecordModificationDate = ""
        RecordModificationTime =  ""

        txt = "Record Key: " + RecordKey + "\r\n"
        txt = txt + "Record Creator Name: "+ RecordCreatorName + "\r\n"
        txt = txt + "Record Creation Date: " + RecordCreationDate + "\r\n"
        txt = txt + "Record Creation Time: " + RecordCreationTime + "\r\n"
        txt = txt + "Record Modifier Name: " + RecordModifierName + "\r\n"
        txt = txt + "Record Modification Date: " + RecordModificationDate + "\r\n"
        txt = txt + "Record Modification Time: " + RecordModificationTime + "\r\n"

        self.m_TextRecordInfo.SetLabelText (txt)
        self.m_TextRecordNum.SetLabelText("")


    def OnSaveRecord(self, event):

        #if read only mode then exit
        if gn.mode_readonly == 'True':  #read-only mode is true, which means no saving permitted
            wx.MessageBox("You cannot make changes to the program while in read-only mode","Save Unsuccessful",wx.OK | wx.ICON_INFORMATION)
            return

        #get the project type ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectType = self.m_ComboProjectType.GetValue()
        #print(projecttypeold)
        if len(ProjectType) <= 1:
             wx.MessageBox("Please select a valid Project Type", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        #get the bid open ("Yes" or "No") ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        BidOpen  = self.m_ComboBidOpen.GetValue()
        if len(BidOpen) <= 1:
             wx.MessageBox("Please select a valid choice for Bid Open", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        # get the status of the project ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        ProjectStatus  = self.m_ComboProjectStatus.GetValue()
        if len(ProjectStatus) <= 1:
             wx.MessageBox("Please select a valid choice for Project Status", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        # get the zone of the project ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Zone= self.m_ComboZone.GetValue()
        if len(Zone) < 1 or Zone.isdigit() == False:
             wx.MessageBox("Please select a valid Project Zone", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return


        # get the quote number ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        QuoteNumber = str(self.m_TextQuoteNumber.GetValue())
        if len(QuoteNumber) < 1 or QuoteNumber.isdigit() == False:
             wx.MessageBox("Please select a valid QuoteNumber", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        # get the quote revision level ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        RevLevel = str(self.m_ComboRevLevel.GetValue())
        if len(RevLevel) < 1 or RevLevel.isdigit() == False:
             wx.MessageBox("Please select a valid Revision level", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return


         # get the sales order number (may be left blank) ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        SalesOrderNum = str(self.m_TextSalesOrderNum.GetValue())
        SalesOrderNum = SalesOrderNum.strip()

        ContinueSave = wx.YES
        if len(SalesOrderNum) > 0 and ProjectType.find("Order") == -1:
            msg = "You have entered a Sales Order number but the record type is " + ProjectType + ", not an Order. "
            msg += "Continue with save and update existing record?"
            ContinueSave = wx.MessageBox(msg, "Save Confirmation",wx.YES_NO | wx.ICON_INFORMATION)

        if ContinueSave == wx.NO:
            return

        # get the drop ship order number (may be left blank) ~~~~~~~~~~~~~~~~~~~~~~~~~~
        DropShipOrderNum= str(self.m_TextDropShipOrderNum.GetValue())

        # get date received ~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx_Rec = self.m_DateReceived.GetValue()
        PyReceived = gn.wxdate2pydate(wx_Rec)                   #Date received in python format

        if wx_Rec.IsValid() == True:
            DateReceived = wx_Rec.FormatISODate()
        else:
            wx.MessageBox("Please select a valid Date Received", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
            return


         # get date requested ~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx_Req = self.m_DateDue.GetValue()

        if wx_Req.IsValid() == True and wx_Rec.IsLaterThan(wx_Req):
            wx.MessageBox("Due Date cannot be earlier than date received", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
            return

        elif wx_Req.IsValid() == True:
            DateRequest = wx_Req.FormatISODate()
        else:
            wx.MessageBox("Please select a valid Date Due", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
            return

        # get date job completed ~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx_Comp = self.m_DateCompleted.GetValue()
        PyComplete = gn.wxdate2pydate(wx_Comp)                        #Date completed in python format

        if wx_Comp.IsValid() == True and wx_Rec.IsLaterThan(wx_Comp) == True:
            wx.MessageBox("Date Completed cannot be less than Date Received", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
            return
        elif wx_Comp.IsValid() == True:
            DateComp = wx_Comp.FormatISODate()
            TADays = str(workdays.networkdays(PyReceived, PyComplete)-1)
            self.m_TextTurnAround.SetLabelText(TADays)
        else:
            DateComp = "1/1/9999"  #Invalid Date
            self.m_TextTurnAround.SetLabelText("TBD")


        # get the name of the appl engineer selected in the combo box ~~~~~~~~~~~
        Assigned = str(self.m_ComboAssigned.GetValue())
        if len(Assigned) <= 1:
             wx.MessageBox("Please select a valid name for Application Engineer", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return


        # get the name of the appl engineer selected in the combo box ~~~~~~~~~~~
        Saleperson = str(self.m_ComboSP.GetValue())
        if len(Saleperson) <= 1:
             wx.MessageBox("Please select a valid name for Saleperson", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return


        # get the Equipment, Buyout, multiplier  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        EquipPrice = str(self.m_TextEquipUSD.GetValue())
        EquipPriceFloat = self.GetFloatValue(EquipPrice)
        if  EquipPriceFloat == -1: EquipPrice = "0"
        else: EquipPrice = str(EquipPriceFloat)

        BuyoutPrice = str(self.m_TextBODollars.GetValue())
        BuyoutPriceFloat = self.GetFloatValue(BuyoutPrice)
        if  BuyoutPriceFloat == -1: BuyoutPrice = "0"
        else: BuyoutPrice = str(BuyoutPriceFloat)

        BuyoutPercent = str(self.m_TextBOPercent.GetValue())
        BuyoutPercentFloat = self.GetFloatValue(BuyoutPercent)
        if  BuyoutPercentFloat == -1: BuyoutPercent = "0"
        else: BuyoutPercent = str(BuyoutPercentFloat)

        Multiplier = str(self.m_TextMultiplier.GetValue())
        MultiplierFloat = self.GetFloatValue(Multiplier)
        if  MultiplierFloat == -1: Multiplier = "0"
        else: Multiplier = str(MultiplierFloat)


        # Get the project folder
        ProjectFolder = str(self.m_TextProjectFolder.GetValue())

        # get the customer name, number, key and ship to address ~~~~~~~
        Customer = str(self.m_ComboCustomerName.GetValue())
        Customer = Customer.replace("&","and")
        if len(Customer) <= 1:
             wx.MessageBox("Please select a valid customer", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        CustomerNumber = str(self.m_TextCustomerNumber.GetValue())
        if len(CustomerNumber) < 1 or CustomerNumber.isdigit() == False:
             wx.MessageBox("Please select a valid customer number", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
             return

        CustomerKey = Customer + " " + CustomerNumber

        ShipTO = str(self.m_TextShipAddress.GetValue())   #may be left blank
        ShipTO = ShipTO.replace("&","and")


        #get notes (may be blank)
        Comments = self.m_TextNotes.GetValue()
        Comments = Comments.replace("&","and")
        Comments = Comments.replace("\n","\r\n")

        #get the alternate refrigerant and CMAT
        CO2 = self.m_CheckCO2.GetValue()
        Ammonia = self.m_CheckAmmonia.GetValue()
        Glycol = self.m_CheckGlycol.GetValue()
        CMAT = self.m_ComboCMAT.GetValue()

        if CO2 == False and Ammonia == False and Glycol == False:
            CMAT = None


        #Get the record information (box in lower right hand corner of the screen) ~~~~~
        # get the record key
        RecordKey = QuoteNumber + ProjectType + RevLevel
        RecordKeyExists = False

        #Record Modification Log
        RecordModifierName = str(gn.user)
        RecordModificationDate = str(wx.DateTime_Today().FormatISODate())
        RecordModificationTime = str(wx.DateTime_Now().FormatISOTime())

        #Check to see if the record key already exists in the database. If it does, then the record creation time
        #should also exist. If a new record is being saved, then the record modification date/time is the same as
        # original record creation time (that is, today's date and the current time)

        sql = "SELECT * FROM dbo.QuoteMaker WHERE RecordKey=\'" + RecordKey + "\'"
        self.dbCursor.execute(sql)
        row = self.dbCursor.fetchone()

        if row != None:
            RecordKeyExists = True

            RecordCreatorName = str(row.RecordCreatorName)
            RecordCreationDate = self.PY2STR_DATE(row.RecordCreationDate)
            RecordCreationTime = self.PY2STR_TIME(row.RecordCreationTime)

        else:
            RecordCreatorName = str(gn.user)
            RecordCreationDate = str(wx.DateTime_Today().FormatISODate())
            RecordCreationTime = str(wx.DateTime_Now().FormatISOTime())

        sqlorder = "SELECT * FROM dbo.QuoteMaker WHERE QuoteNumber = '" + QuoteNumber + "'and projecttype = 'Order'"
        self.dbCursor.execute(sqlorder)
        row = self.dbCursor.fetchone()
        if row != None:
            OrderExists = True
        else:
            OrderExists = False

        # If the record does not exisit, check to make sure that AE, Saleperson and Customer are in the allowed list
        if RecordKeyExists == False:
                try: self.AENames.index(Assigned)
                except:
                        wx.MessageBox("The name of Application was not found in the allowed list", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                        return

                try: self.SPNames.index(Saleperson)
                except:
                        wx.MessageBox("The name of Saleperson was not found in the allowed list", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                        return

                try: self.CustomerData.index(Customer)
                except:
                        wx.MessageBox("The name of the Customer was not found in the allowed list", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                        return

                try: self.CustomerNumbers.index(int(CustomerNumber))
                except:
                        wx.MessageBox("The Customer Number was not found in the allowed list", "Save Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
                        return



        # If the record exists, check if the user wants to over-write the record
        ContinueSave = wx.YES
        if RecordKeyExists == True:
            msg = ProjectType + " " + QuoteNumber  + " Revision " + RevLevel + " already exists. Continue with save and update existing record?"
            ContinueSave = wx.MessageBox(msg, "Save Confirmation",wx.YES_NO | wx.ICON_INFORMATION)

        if ContinueSave == wx.NO:
            return

        #Check to see if there is ANOTHER record in database using the same quote number (but a different customer number).
        # All records that share the same quote number should ideally have the same customer number as well
        sql = "SELECT * FROM dbo.QuoteMaker WHERE QuoteNumber=" + str(QuoteNumber)
        self.dbCursor.execute(sql)
        SingleRow = self.dbCursor.fetchone()

        ContinueSave = wx.YES
        if SingleRow != None and RecordKeyExists == False:
            if int(CustomerNumber) != SingleRow.CustomerNumber:
                msg = "There is an existing record in the database with \r\n\r\n Quote Number: " + str(QuoteNumber) + \
                "\r\n Project Type: " + str(SingleRow.ProjectType) + "\r\n Customer Number: " + str(SingleRow.CustomerNumber) + "\r\n\r\n" + \
                "You are trying to save a record under the same quote number but a different Customer Number. Would you like to continue" \
                " with the saving the record? It is recommended you select NO, and check your customer information again."
                ContinueSave = wx.MessageBox(msg, "Save Confirmation",wx.YES_NO | wx.ICON_INFORMATION)

        if ContinueSave == wx.NO:
            return


        try:
            prev_ProjectType, prev_DateComp, prev_SalesOrderNum = self.dbCursor.execute("SELECT ProjectType, DateComp, SalesOrderNum FROM dbo.QuoteMaker WHERE RecordKey=?", self.DBRecordKeys[self.Index]).fetchone()
            
            prev_DateComp = str(prev_DateComp).replace(' 00:00:00', '')
            
            current_ProjectType = ProjectType
            current_SalesOrderNum = SalesOrderNum
            
            if wx_Comp.IsValid():
                current_DateComp = wx_Comp.FormatISODate()
            else:
                current_DateComp = '9999-01-01'
            
            #print('   prev_ProjectType: {}'.format(prev_ProjectType))
            #print('current_ProjectType: {}'.format(current_ProjectType))
            #print('   prev_DateComp: {}'.format(prev_DateComp))
            #print('current_DateComp: {}'.format(current_DateComp))
            #print('   prev_SalesOrderNum: {}'.format(prev_SalesOrderNum))
            #print('current_SalesOrderNum: {}'.format(current_SalesOrderNum))
        
            #check id project type old and new are different to send email
            if prev_ProjectType != current_ProjectType and current_ProjectType == 'Order' and OrderExists == False:
                print("ProjectType has changed to Order.")
                self.sendemail('order',QuoteNumber,RevLevel,Assigned);

            #check id completion date old and new are different to send email
            if prev_DateComp != current_DateComp and current_DateComp != '9999-01-01' and OrderExists == True:
                print("Completion date has changed.")
                self.sendemail('compdate',QuoteNumber,RevLevel,RecordCreatorName);

            #check id salesorder and dropshipto number old and new are different to send email
            if prev_SalesOrderNum != current_SalesOrderNum and len(current_SalesOrderNum) != 0 and OrderExists == True:
                print("SalesOrder has changed.")
                self.sendemail('salesorder',QuoteNumber,RevLevel,Assigned);
        
        except Exception as e:
            print(e)


        ######################### CREATE THE SQL QUERY TO SAVE ##################################
        sql = "INSERT INTO dbo.QuoteMaker (ProjectType, BidOpen, ProjectStatus, Zone, QuoteNumber, RevLevel," \
              "SalesOrderNum, DropShipOrderNum, DateReceived, DateRequest, DateComp, Assigned, Saleperson, EquipPrice, " \
              "BuyoutPrice, BuyoutPercent, Multiplier, ProjectFolder, Customer, CustomerKey, CustomerNumber, ShipTO, " \
              "RecordKey, RecordCreatorName, RecordCreationDate, RecordCreationTime, RecordModifierName, RecordModificationDate, " \
              "RecordModificationTime, Comments, CarbonDioxide, Ammonia, Glycol, CMAT) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?" \
              ",?,?,?,?,?,?,?,?,?)"


        sql_update = "UPDATE dbo.QuoteMaker SET ProjectType=?, BidOpen=?, ProjectStatus=?, Zone=?,QuoteNumber=?, " \
                     "RevLevel=?, SalesOrderNum=?, DropShipOrderNum=?, DateReceived=?, DateRequest=?, DateComp=?, " \
                     "Assigned=?,Saleperson=?,EquipPrice=?,BuyoutPrice=?, BuyoutPercent=?, Multiplier=?, " \
                     "ProjectFolder=?, Customer=?,CustomerKey=?,CustomerNumber=?,ShipTO=?,RecordCreatorName=?," \
                     "RecordCreationDate=?, RecordCreationTime=?, RecordModifierName=?, RecordModificationDate=?, " \
                     "RecordModificationTime=?,Comments=?, CarbonDioxide=?,Ammonia=?,Glycol=?,CMAT=? WHERE RecordKey=?"


        if RecordKeyExists == True:
            self.dbCursor.execute(sql_update,ProjectType, BidOpen,ProjectStatus,int(Zone),int(QuoteNumber), int(RevLevel),
                                  SalesOrderNum,DropShipOrderNum,DateReceived, DateRequest, DateComp, Assigned, Saleperson,
                                  float(EquipPrice), float(BuyoutPrice), float(BuyoutPercent), float(Multiplier),
                                  ProjectFolder,Customer, CustomerKey, int(CustomerNumber),ShipTO,RecordCreatorName,
                                  RecordCreationDate,  RecordCreationTime, RecordModifierName, RecordModificationDate,
                                  RecordModificationTime,  Comments,CO2, Ammonia, Glycol, CMAT, RecordKey)
            self.dbCursor.commit()
            wx.MessageBox("The record was successfully saved","Save Successful",wx.OK | wx.ICON_INFORMATION)
            return



        #Record key does not exisit so a new record will be created
        self.dbCursor.execute(sql,ProjectType, BidOpen, ProjectStatus,int(Zone),int(QuoteNumber), int(RevLevel),
                                    SalesOrderNum, DropShipOrderNum,DateReceived, DateRequest, DateComp,
                                    Assigned, Saleperson,  float(EquipPrice), float(BuyoutPrice), float(BuyoutPercent),
                                    float(Multiplier), ProjectFolder,Customer, CustomerKey, int(CustomerNumber), ShipTO,
                                    RecordKey, RecordCreatorName, RecordCreationDate,  RecordCreationTime, RecordModifierName,
                                    RecordModificationDate,  RecordModificationTime,  Comments,CO2, Ammonia, Glycol, CMAT)
        self.dbCursor.commit()

        #Add the new record to the DBRecordKeys list
        self.DBRecordKeys.append(RecordKey)
        self.count = self.count + 1    #increase the record count
        self.Index = self.count - 1    #set the index to the last record
        self.LoadRecord()              #load the (newly added) last record to the display

        wx.MessageBox("The record was successfully saved","Save Successful",wx.OK | wx.ICON_INFORMATION)


    def OnComboCustomerSelected(self, event):

                name = self.m_ComboCustomerName.GetValue()

                #find the customer in the list
                try:  idx = self.CustomerData.index(name)
                except: idx = -1

                if idx < 0: number = ""   # customer not found in the list
                else: number = str(self.CustomerNumbers[idx])

                key = name + " " + number

                self.m_TextCustomerNumber.SetLabelText(number)
                self.m_TextCustomerKey.SetLabelText(key)

    def OnCheckRefrigerant(self, event):

        CO2 = self.m_CheckCO2.GetValue()
        Ammonia = self.m_CheckAmmonia.GetValue()
        Glycol = self.m_CheckGlycol.GetValue()

        if CO2 == True or Ammonia == True or Glycol == True:
            self.ShowCMAT()
        else:
            self.HideCMAT()

    def GetFloatValue(self, str_value):

                str_value = str_value.replace("$","")
                str_value = str_value.replace(",","")
                str_value = str_value.strip()

                if len(str_value) == 0:
                                return -1
                try:
                                float_value = float(str_value)
                                return float_value
                except ValueError:
                                return -1


    def PY2STR_DATE(self,pyDateTime):

        if pyDateTime == None: return ""

        try:
            dt = pyDateTime.date()
            dt_str = str(dt)
            return dt_str

        except ValueError:
            return ""

    def PY2STR_TIME(self,pyDateTime):

        if pyDateTime == None: return ""

        try:
            pyDateTime = pyDateTime.time()
            dt_str = pyDateTime.strftime("%I:%M %p")
            return dt_str

        except ValueError:
            return ""


    def OnNightlyBuild(self):
        args = [gn.NightlyBuildPath,'/SILENT']
        subprocess.call(args)


    def OnExitApp( self, event ):
        self.mainWnd.Close()


    def OnManageUsers(self,event):

        #Check the delete pass phrase
        notesText = self.m_TextNotes.GetValue()
        if notesText.find("superadmin") == -1:
            MsgWnd = MessageDialog(self)
            MsgWnd.ShowModal()
            return

        ManageUsersWnd = UserManagerDialog(self)
        ManageUsersWnd.dbCur = self.dbCursor
        ManageUsersWnd.LoadList()
        ManageUsersWnd.ShowModal()


    def OnAboutBox(self, event):

        info = wx.AboutDialogInfo()
        ImgPath = str(os.path.split(__file__)[0]) + "\Images\AppIcon.png"
        info.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))
        info.SetName('Quote Manager')
        info.SetVersion('1.00')
        info.SetCopyright('(C) 2015 Heatcraft Worldwide Refrigeration')
        info.SetWebSite('http://www.heatcraftrpd.com')
        info.SetDescription('Phone: 706-568-1514')
        wx.AboutBox(info)


    def PepareHTMLText(self):

        ProjectType = self.m_ComboProjectType.GetValue()
        BidOpen  = self.m_ComboBidOpen.GetValue()
        ProjectStatus  = self.m_ComboProjectStatus.GetValue()
        Zone= self.m_ComboZone.GetValue()

        QuoteNumber = str(self.m_TextQuoteNumber.GetValue())
        RevLevel = str(self.m_ComboRevLevel.GetValue())
        SalesOrderNum = str(self.m_TextSalesOrderNum.GetValue())
        DropShipOrderNum= str(self.m_TextDropShipOrderNum.GetValue())

        wx_Rec = self.m_DateReceived.GetValue()
        if wx_Rec.IsValid() == True: DateReceived = wx_Rec.FormatISODate()
        else: DateReceived = ""

        wx_Req = self.m_DateDue.GetValue()
        if wx_Req.IsValid() == True: DateRequest = wx_Req.FormatISODate()
        else:DateRequest = ""

        wx_Comp = self.m_DateCompleted.GetValue()
        if wx_Comp.IsValid() == True:DateComp = wx_Comp.FormatISODate()
        else:DateComp = ""

        TADays = self.m_TextTurnAround.GetValue()

        Assigned = str(self.m_ComboAssigned.GetValue())
        Saleperson = str(self.m_ComboSP.GetValue())

        EquipPrice = str(self.m_TextEquipUSD.GetValue())
        BuyoutPrice = str(self.m_TextBODollars.GetValue())
        BuyoutPercent = str(self.m_TextBOPercent.GetValue())
        Multiplier = str(self.m_TextMultiplier.GetValue())

        ProjectFolder = str(self.m_TextProjectFolder.GetValue())
        Customer = str(self.m_ComboCustomerName.GetValue())
        CustomerNumber = str(self.m_TextCustomerNumber.GetValue())
        CustomerKey = str(self.m_TextCustomerKey.GetValue())
        ShipTO = str(self.m_TextShipAddress.GetValue())

        Comments = str(self.m_TextNotes.GetValue()).strip()
        Comments = Comments.replace("\n","<br>")

        RecordText = self.m_TextRecordInfo.GetValue()
        RecordText = RecordText.replace("\n","<br>")

        ######################### CREATE THE HTML TABLE ##################################

        html = '<table>'
        html += '''<tr><td align="left" valign="top" nowrap> </td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap><strong>PROJECT INFORMATION</strong></td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap> </td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap> ProjectType: {}&nbsp;</td></tr>'''.format(ProjectType)
        html += '''<tr><td align="left" valign="top" nowrap>BidOpen: {}&nbsp;</td></tr>'''.format(BidOpen)
        html += '''<tr><td align="left" valign="top" nowrap>ProjectStatus: {}&nbsp;</td></tr>'''.format(ProjectStatus)
        html += '''<tr><td align="left" valign="top" nowrap>Zone: {}&nbsp;</td></tr>'''.format(Zone)
        html += '''<tr><td align="left" valign="top" nowrap>QuoteNumber: {}&nbsp;</td></tr>'''.format(QuoteNumber)
        html += '''<tr><td align="left" valign="top" nowrap>RevLevel: {}&nbsp;</td></tr>'''.format(RevLevel)
        html += '''<tr><td align="left" valign="top" nowrap>SalesOrderNum: {}&nbsp;</td></tr>'''.format(SalesOrderNum)
        html += '''<tr><td align="left" valign="top" nowrap>DropShipOrderNum: {}&nbsp;</td></tr>'''.format(DropShipOrderNum)
        html += '''<tr><td align="left" valign="top" nowrap>DateReceived: {}&nbsp;</td></tr>'''.format(DateReceived)
        html += '''<tr><td align="left" valign="top" nowrap>DateRequest: {}&nbsp;</td></tr>'''.format(DateRequest)
        html += '''<tr><td align="left" valign="top" nowrap>DateComp: {}&nbsp;</td></tr>'''.format(DateComp)
        html += '''<tr><td align="left" valign="top" nowrap>TADays: {}&nbsp;</td></tr>'''.format(TADays)
        html += '''<tr><td align="left" valign="top" nowrap>Assigned: {}&nbsp;</td></tr>'''.format(Assigned)
        html += '''<tr><td align="left" valign="top" nowrap>Saleperson: {}&nbsp;</td></tr>'''.format(Saleperson)
        html += '''<tr><td align="left" valign="top" nowrap>EquipPrice: {}&nbsp;</td></tr>'''.format(EquipPrice)
        html += '''<tr><td align="left" valign="top" nowrap>BuyoutPrice: {}&nbsp;</td></tr>'''.format(BuyoutPrice)
        html += '''<tr><td align="left" valign="top" nowrap>BuyoutPercent: {}&nbsp;</td></tr>'''.format(BuyoutPercent)
        html += '''<tr><td align="left" valign="top" nowrap>Multiplier: {}&nbsp;</td></tr>'''.format(Multiplier)

        html += '''<tr><td align="left" valign="top" nowrap> </td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap><strong>CUSTOMER INFORMATION</strong></td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap> </td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap>Customer: {}&nbsp;</td></tr>'''.format(Customer)
        html += '''<tr><td align="left" valign="top" nowrap>CustomerKey: {}&nbsp;</td></tr>'''.format(CustomerKey)
        html += '''<tr><td align="left" valign="top" nowrap>CustomerNumber: {}&nbsp;</td></tr>'''.format(CustomerNumber)
        html += '''<tr><td align="left" valign="top" nowrap>ShipTO: {}&nbsp;</td></tr>'''.format(ShipTO)

        html += '''<tr><td align="left" valign="top" nowrap> </td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap><strong>RECORD INFORMATION</strong></td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap> </td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap>{}&nbsp;</td></tr>'''.format(RecordText)
        html += '''<tr><td align="left" valign="top" nowrap><strong>COMMENTS</strong></td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap>{}&nbsp;</td></tr>'''.format(Comments)
        html += '</table>'

        self.PrintText = html

    def PepareHTMLTextforemail(self):

        #prepare html for sending email
        ProjectType = self.m_ComboProjectType.GetValue()
        BidOpen  = self.m_ComboBidOpen.GetValue()
        ProjectStatus  = self.m_ComboProjectStatus.GetValue()
        Zone= self.m_ComboZone.GetValue()

        QuoteNumber = str(self.m_TextQuoteNumber.GetValue())
        RevLevel = str(self.m_ComboRevLevel.GetValue())
        SalesOrderNum = str(self.m_TextSalesOrderNum.GetValue())
        DropShipOrderNum= str(self.m_TextDropShipOrderNum.GetValue())

        wx_Rec = self.m_DateReceived.GetValue()
        if wx_Rec.IsValid() == True: DateReceived = wx_Rec.FormatISODate()
        else: DateReceived = ""

        wx_Req = self.m_DateDue.GetValue()
        if wx_Req.IsValid() == True: DateRequest = wx_Req.FormatISODate()
        else:DateRequest = ""

        wx_Comp = self.m_DateCompleted.GetValue()
        if wx_Comp.IsValid() == True:DateComp = wx_Comp.FormatISODate()
        else:DateComp = ""

        TADays = self.m_TextTurnAround.GetValue()

        Assigned = str(self.m_ComboAssigned.GetValue())
        Saleperson = str(self.m_ComboSP.GetValue())

        EquipPrice = str(self.m_TextEquipUSD.GetValue())
        BuyoutPrice = str(self.m_TextBODollars.GetValue())
        BuyoutPercent = str(self.m_TextBOPercent.GetValue())
        Multiplier = str(self.m_TextMultiplier.GetValue())

        ProjectFolder = str(self.m_TextProjectFolder.GetValue())
        Customer = str(self.m_ComboCustomerName.GetValue())
        CustomerNumber = str(self.m_TextCustomerNumber.GetValue())
        CustomerKey = str(self.m_TextCustomerKey.GetValue())
        ShipTO = str(self.m_TextShipAddress.GetValue())

        Comments = str(self.m_TextNotes.GetValue()).strip()
        Comments = Comments.replace("\n","<br>")

        RecordText = self.m_TextRecordInfo.GetValue()
        RecordText = RecordText.split('\n')

        #RecordText = [R.replace("\n","<tr><td align="+"""left""" +"valign="+"""top"""+ "nowrap>") for R in RecordText]
        RecordText = [R.replace(":","</td><td>",1) for R in RecordText]
        RecordText = '</tr><tr><td align="left" valign="top" nowrap>'.join(filter(None,RecordText))


        ######################### CREATE THE HTML TABLE #################################

        html = '<table border =1 >'

        html += '''<tr><td align="left" valign="top"  colspan="2" nowrap><strong>PROJECT INFORMATION</strong></td></tr>'''
        #html += '''<tr><td align="left" valign="top" nowrap> </td><td></td></tr>'''
        html += '''<tr><td align="left" valign="top" nowrap> ProjectType:</td> ''' + '''<td>{}&nbsp;</td></tr>'''.format(ProjectType)
        html += '''<tr><td align="left" valign="top" nowrap>QuoteNumber: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(QuoteNumber)
        html += '''<tr><td align="left" valign="top" nowrap>DateReceived: </td>''' + '''<td>{}&nbsp;</td></tr>'''.format(DateReceived)
        html += '''<tr><td align="left" valign="top" nowrap>Salesperson:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(Saleperson)
        html += '''<tr><td align="left" valign="top" nowrap>BuyoutPrice:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(BuyoutPrice)
        html += '''<tr><td align="left" valign="top" nowrap>BidOpen:</td>'''+ '''<td>{}&nbsp;</td></tr>'''.format(BidOpen)
        html += '''<tr><td align="left" valign="top" nowrap>ProjectStatus:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(ProjectStatus)
        html += '''<tr><td align="left" valign="top" nowrap>Zone:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(Zone)
        html += '''<tr><td align="left" valign="top" nowrap>RevLevel: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(RevLevel)
        html += '''<tr><td align="left" valign="top" nowrap>SalesOrder#: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(SalesOrderNum)
        html += '''<tr><td align="left" valign="top" nowrap>DropShipOrder#: </td>'''+'''<td> {}&nbsp;</td></tr>'''.format(DropShipOrderNum)
        html += '''<tr><td align="left" valign="top" nowrap>DateRequest: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(DateRequest)
        html += '''<tr><td align="left" valign="top" nowrap>DateComp: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(DateComp)
        html += '''<tr><td align="left" valign="top" nowrap>TADays: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(TADays)
        html += '''<tr><td align="left" valign="top" nowrap>Assigned: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(Assigned)
        html += '''<tr><td align="left" valign="top" nowrap>EquipPrice:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(EquipPrice)
        html += '''<tr><td align="left" valign="top" nowrap>BuyoutPercent:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(BuyoutPercent)
        html += '''<tr><td align="left" valign="top" nowrap>Multiplier:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(Multiplier)

        html += '''<tr><td align="left" valign="top"  colspan="2" nowrap><strong>CUSTOMER INFORMATION</strong></td></tr>'''

        html += '''<tr><td align="left" valign="top" nowrap>Customer: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(Customer)
        html += '''<tr><td align="left" valign="top" nowrap>CustomerKey:</td>'''+'''<td> {}&nbsp;</td></tr>'''.format(CustomerKey)
        html += '''<tr><td align="left" valign="top" nowrap>ShipTO: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(ShipTO)
        html += '''<tr><td align="left" valign="top" nowrap>CustomerNumber: </td>'''+'''<td>{}&nbsp;</td></tr>'''.format(CustomerNumber)



        html += '''<tr><td align="left" valign="top"  colspan="2" nowrap><strong>RECORD INFORMATION</strong></td></tr>'''

        html += '''<tr><td align="left" valign="top" nowrap>{}&nbsp;</td></tr>'''.format(RecordText)
        html += '''<tr><td align="left" valign="top"  colspan="2" nowrap><strong>COMMENTS</strong></td></tr>'''
        html += '''<tr><td align="left" valign="top" colspan="2" nowrap>{}&nbsp;</td></tr>'''.format(Comments)
        html += '</table>'

        self.PrintText = html
        return html

    def sendemail(self,typeofmail,quote,revision, send_to_name):
        TO = []
        CC = []
        ##Emailnames = []
        ##Emailnames = self.m_checkListBox.GetCheckedStrings();

        ##email_cc_sql = "select emailid from dbo.emailforquote where etype = 'optional' and enabled = 1 and name in ('" + "','".join(Emailnames) +"')"

        #toaddr = "nivedha.dhakshnamurthy@lennoxintl.com"
        msg = MIMEMultipart()
        #mail part for changing qoute to order
        if typeofmail == 'order' :
            email_to_sql = "select top 1 email from dbo.employees where name like '"+ send_to_name +"%'"


            msg['Subject'] = "Quote  :"+quote+" revision : "+revision+" changed to order by "+str(gn.user)

        #mail part for changing completion date
        elif typeofmail == 'compdate':
            email_to_sql = "select top 1 email from dbo.employees where name like '"+ send_to_name +"%'"

            wx_Comp = self.m_DateCompleted.GetValue()
            if wx_Comp.IsValid() == True:DateComp = wx_Comp.FormatISODate()
            else:
                DateComp = ""

            msg['Subject'] = "Quote  :"+quote+" revision : "+revision+" Completed date updated by "+str(gn.user)

        #mail part for changing salesorder and dropshipto number
        elif typeofmail == 'salesorder':
            email_to_sql = "select emailid from emailforquote where emailid is not null and etype = 'order' and enabled=1"

            msg['Subject'] = "Quote  :"+quote+" revision : "+revision+" Salesorder and dropshipto updated "+str(gn.user)


        try:
            db_connection = db.connect_to_eng04_database()
            cursor = db_connection.cursor()
            
            for row in cursor.execute(email_to_sql).fetchall():
                TO.append(row[0])

            if typeofmail == 'compdate':
                #use customer service's respective POD email addresses for CC
                pod_emails = self.dbCursor.execute("SELECT pod_email FROM dbo.customer_service_pods WHERE customer_service_name IN (SELECT RecordCreatorName FROM dbo.QuoteMaker WHERE RecordKey=?)", self.DBRecordKeys[self.Index]).fetchall()

                #if no customer service POD relation is found, use all known POD email addresses for CC
                if not pod_emails:
                    pod_emails = self.dbCursor.execute("SELECT DISTINCT pod_email FROM dbo.customer_service_pods").fetchall()
                    customer_service_name = self.dbCursor.execute("SELECT RecordCreatorName FROM dbo.QuoteMaker WHERE RecordKey=?", self.DBRecordKeys[self.Index]).fetchone()[0]
                    wx.MessageBox("There are no PODs associated with Customer Service user {}.\nEmails will CC all PODs for now.\n\nAsk your Database Administrator to add {}'s respective PODs to dbo.customer_service_pods.".format(customer_service_name, customer_service_name), 'Notice!', wx.OK | wx.ICON_INFORMATION)
            else:
                pod_emails = []

            for row in pod_emails:
                CC.append(row[0])
            
            fromaddr = cursor.execute("select top 1 email from dbo.employees where name = ?", gn.user).fetchone()[0]

            db_connection.close()

            toaddr = ",".join(TO)
            ccaddr = ",".join(CC)

            #fromaddr = "donotreply@lennoxintl.com"
            
            msg['From'] = fromaddr
            msg['To'] = toaddr
            msg['CC'] = ccaddr
            text = "HI , Please check the detail of the quote."
            html = self.PepareHTMLTextforemail();

            print('Sending email:')
            print('   FROM: {}'.format(fromaddr))
            print('   TO: {}'.format(toaddr))
            print('   CC: {}'.format(ccaddr))
            
            part1 = MIMEText(text, 'plain')
            part2 = MIMEText(html, 'html')
            msg.attach(part2)
            server = smtplib.SMTP('mailrelay.lennoxintl.com',25)
            tex = msg.as_string()
            server.sendmail(fromaddr,TO+CC,tex)
            server.quit()
            ProjectType = self.m_ComboProjectType.GetValue()
            SalesOrderNum = str(self.m_TextSalesOrderNum.GetValue())
            DropShipOrderNum= str(self.m_TextDropShipOrderNum.GetValue())
            if typeofmail == 'salesorder' :
                global DropShipOrderNumberold
                DropShipOrderNumberold = DropShipOrderNum
                global salesorderold
                salesorderold = SalesOrderNum
            elif typeofmail == 'compdate':
                global datejobcompletedold
                datejobcompletedold = str(wx_Comp)
            elif typeofmail == 'order':
                global projecttypeold
                projecttypeold = ProjectType
        except Exception as e:
            wx.MessageBox("Record will be saved and Mail must be sent manually","Error Sending Mail",wx.OK | wx.ICON_INFORMATION)
            pass

