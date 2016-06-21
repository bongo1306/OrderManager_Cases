import wx
from wx import xrc
ctrl = xrc.XRCCTRL
import os


#####################################################################################################################
#####################################################################################################################
class UserManagerDialog(wx.Dialog):
	def __init__(self, parent):
		#load dialog XRC description
		pre = wx.PreDialog()
		res = xrc.XmlResource.Get()
		res.LoadOnDialog(pre, parent, "UserManagerDialog")
		self.PostCreate(pre)
		self.SetSize((500,600))
		ImgPath = str(os.path.split(__file__)[0]) + "\Images\Lists.png"
		self.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))
		self.dbCur = None
		self.SortCriteria = " ORDER BY Name"


		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.Bind(wx.EVT_COMBOBOX, self.OnChangeList)    #user picked a different list to manage
		self.Bind(wx.EVT_BUTTON, self.OnAddUser, id=xrc.XRCID('m_BtnAddUser'))      # Add user to the currently displayed list
		self.Bind(wx.EVT_BUTTON, self.OnDeleteUser, id=xrc.XRCID('m_BtnDeleteUser'))   #Delete user from currently displayed list
		self.Bind(wx.EVT_LIST_COL_CLICK, self.OnHeaderSort,id=xrc.XRCID('m_ListManageUsers') )


		self.m_ListManageUsers = wx.FindWindowByName('m_ListManageUsers')
		self.m_ComboManageUsers = wx.FindWindowByName('m_ComboManageUsers')
		self.m_BtnAddUser = wx.FindWindowByName('m_BtnAddUser')
		self.m_BtnDeleteUser = wx.FindWindowByName('m_BtnDeleteUser')

		self.m_ListManageUsers.InsertColumn(0, "#" )
		self.m_ListManageUsers.InsertColumn(1, "Name" )

		self.m_ComboManageUsers.Select(0)


	def on_close_frame(self, event):
		self.Destroy()

    # User chooses another list to manage
	def OnChangeList(self, event):
		self.LoadList()


    # User chooses another list to manage
	def LoadList(self):

		listName = self.m_ComboManageUsers.GetValue()
		listName = listName.lower()

		sql = ""
		if listName.find("appl") != wx.NOT_FOUND :	 sql = "SELECT * FROM dbo.ApplicationsTable" + self.SortCriteria
		elif listName.find("cust") != wx.NOT_FOUND : sql = "SELECT * FROM dbo.CustomerTable" + self.SortCriteria
		elif listName.find("sale") != wx.NOT_FOUND : sql = "SELECT * FROM dbo.SalepersonTable" + self.SortCriteria

		Record = self.dbCur.execute(sql).fetchall()
		NumRes = len(Record)
		self.m_ListManageUsers.DeleteAllItems()
		if NumRes == 0:	return  #No search results found
		x = 0

		while x < NumRes:
				if Record[x].Number == None: Num = "-1"
				else: Num = str(Record[x].Number)

				if Record[x].Name == None: Name = "NA"
				else: Name = str(Record[x].Name)

				self.m_ListManageUsers.InsertStringItem(x,Num)
				self.m_ListManageUsers.SetStringItem(x,1, Name)
				x = x + 1

		self.m_ListManageUsers.SetColumnWidth(0, 75)
		self.m_ListManageUsers.SetColumnWidth(1, wx.LIST_AUTOSIZE_USEHEADER)


    # Add new user
	def OnAddUser(self, event):

		listName = self.m_ComboManageUsers.GetValue()
		listName = listName.lower()
		dlg = AddUserDialog(self)
		dlg.CursorName = self.dbCur

		if listName.find("appl") != wx.NOT_FOUND :	 dlg.tableName = "dbo.ApplicationsTable"
		elif listName.find("cust") != wx.NOT_FOUND : dlg.tableName = "dbo.CustomerTable"
		elif listName.find("sale") != wx.NOT_FOUND : dlg.tableName = "dbo.SalepersonTable"

		dlg.SetUserNumber()
		dlg.ShowModal()



    # Delete a user from list
	def OnDeleteUser(self, event):
        #Get the name of the table from which a record will be deleted
		listName = self.m_ComboManageUsers.GetValue()
		listName = listName.lower()
		dbTable = ""

		if listName.find("appl") != wx.NOT_FOUND :	 dbTable = "dbo.ApplicationsTable"
		elif listName.find("cust") != wx.NOT_FOUND : dbTable = "dbo.CustomerTable"
		elif listName.find("sale") != wx.NOT_FOUND : dbTable = "dbo.SalepersonTable"

        #Get the record that needs to be deleted
		index = self.m_ListManageUsers.GetFirstSelected()

		if index < 0:
				wx.MessageBox("No valid record selected for deletion", "Deletion Unsuccessful",wx.OK | wx.ICON_EXCLAMATION)
				return

		NumberStr = self.m_ListManageUsers.GetItem(index, 0).GetText()
		NameStr = self.m_ListManageUsers.GetItem(index, 1).GetText()
		sql = "DELETE FROM " + dbTable + " WHERE Number=" + NumberStr + " AND Name=\'" + NameStr + "\'"
		msg = "Would you like to delete " + NameStr + " from " + dbTable + "? "

		#Confirm before deleting record
		ConfirmDelete = wx.MessageBox(msg, "Save Confirmation",wx.YES_NO | wx.ICON_INFORMATION)
		if ConfirmDelete == wx.NO: return

		self.dbCur.execute(sql)
		self.dbCur.commit()
		self.LoadList()

        #Update the combo boxes in the main window and reload
		pWnd = self.GetParent()
		if listName.find("appl") != wx.NOT_FOUND :	 pWnd.FillAEComboBoxNames()
		elif listName.find("sale") != wx.NOT_FOUND : pWnd.FillSPComboBoxNames()
		elif listName.find("cust") != wx.NOT_FOUND : pWnd.FillCustomerComboBox()

		pWnd.LoadRecord()               #update main window


   # Sort the results according to the column header clicked by the user
	def OnHeaderSort(self, event):
		column = event.GetColumn()
		if column == 0: self.SortCriteria = self.SortCriteria = " ORDER BY Number"
		elif column == 1: self.SortCriteria = self.SortCriteria = " ORDER BY Name"
		self.LoadList()




#####################################################################################################################
#####################################################################################################################
class AddUserDialog(wx.Dialog):
	def __init__(self, parent):
		#load frame XRC description
		pre = wx.PreDialog()
		res = xrc.XmlResource.Get()
		res.LoadOnDialog(pre, parent, "AddUserDialog")
		self.PostCreate(pre)
		ImgPath = str(os.path.split(__file__)[0]) + "\Images\Lists.png"
		self.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))
		self.tableName = ""
		self.CursorName = None

		#bindings
		self.Bind(wx.EVT_CLOSE, self.on_close_frame)
		self.Bind(wx.EVT_BUTTON, self.OnOK, id=xrc.XRCID('m_BtnAddOK'))
		self.Bind(wx.EVT_BUTTON, self.OnCancel, id=xrc.XRCID('m_BtnAddCancel'))

		self.m_TextAddUserNum = wx.FindWindowByName('m_TextAddUserNum')
		self.m_TextAddUserName = wx.FindWindowByName('m_TextAddUserName')


	#The user number is not needed for application engineers and sales list
	def SetUserNumber(self):

		if self.tableName.find("dbo.CustomerTable") != -1:
			return  # if this is customer list, then return.

		self.m_TextAddUserNum.SetEditable(False)    #Disable the control

		# check to see if the customer number already exists
		sql = "SELECT MAX(Number) AS HighestNum FROM " + self.tableName;
		self.CursorName.execute(sql)
		row = self.CursorName.fetchone()
		if row != None: self.m_TextAddUserNum.SetLabelText(str(row.HighestNum+1))


	def on_close_frame(self, event):
		self.Destroy()

	def OnOK(self, event):
		NumberStr = str(self.m_TextAddUserNum.GetValue())
		NameStr =   str(self.m_TextAddUserName.GetValue())

		if len(NumberStr) < 1  or NumberStr.isdigit() == False:
				wx.MessageBox("Please enter a valid number")
				return

		if len(NameStr) <= 1:
				wx.MessageBox("Please enter a valid name")
				return

		#Discourage the use of ampersand character in customer and people names.
		NameStr = NameStr.replace("&","and")

		# check to see if the customer name already exists
		sql = "SELECT * FROM " + self.tableName + " WHERE Name=?"
		self.CursorName.execute(sql, NameStr)
		row = self.CursorName.fetchone()

		if row != None:
				wx.MessageBox("The customer name already exists. Please enter a unique name for the customer")
				return

		# check to see if the customer number already exists
		sql = "SELECT * FROM " + self.tableName + " WHERE Number=" + NumberStr
		self.CursorName.execute(sql)
		row = self.CursorName.fetchone()

		if row != None:
				wx.MessageBox("The customer number already exists. Please enter a unique name for the customer")
				return


		sql = "INSERT INTO " + self.tableName + " (Number, Name) VALUES (" + NumberStr + ",\'" + NameStr + "\')"
		self.CursorName.execute(sql)
		self.CursorName.commit()

		pUserManagerWnd = self.GetParent()
		pUserManagerWnd.LoadList()          #update new user in the user manager list

		pMainWnd = pUserManagerWnd.GetParent()
		if self.tableName.find("Appl") != wx.NOT_FOUND :   pMainWnd.FillAEComboBoxNames()
		elif self.tableName.find("Sale") != wx.NOT_FOUND : pMainWnd.FillSPComboBoxNames()
		elif self.tableName.find("Cust") != wx.NOT_FOUND : pMainWnd.FillCustomerComboBox()
		pMainWnd.LoadRecord()               #update main window

		self.Destroy()

	def OnCancel(self, event):
		self.Destroy()