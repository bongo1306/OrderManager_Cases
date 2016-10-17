import wx
#import wx.xrc
from wx import xrc
ctrl = xrc.XRCCTRL
import wx.calendar
import wx.lib.inspection
from win32com.client import Dispatch
from dateutil.parser import parse
#import numpy as np
#import easygui
import pypyodbc
import re
import datetime
import General as gn
import Database as db
import threading
import sys
import os
import csv
import os
import time
import re

class AdminTab(object):

        def init_admin_tab(self):                               
                              
                # Connect Events
                 
                self.name=wx.FindWindowByName('choice:name')
                self.notebook =wx.FindWindowByName('notebook:applications')
                #self.m_panel9=wx.FindWindowByName('m_panel9')
                #self.m_panel9 = wx.Panel( self.applications)
                self.createuserpanel=wx.FindWindowByName('createuserpanel')
                self.m_staticName = wx.FindWindowByName('m_staticName')
                self.m_textName = wx.FindWindowByName('m_textName')   
                self.m_staticdepartment = wx.FindWindowByName('m_staticdepartment')
                self.m_combodepartment = wx.FindWindowByName('m_combodepartment')
                self.m_staticEmail = wx.FindWindowByName('m_staticEmail')
                self.m_textEmail = wx.FindWindowByName('m_textEmail')
                self.m_staticPassword = wx.FindWindowByName('m_staticPassword')
                self.m_textpassword = wx.FindWindowByName('m_textpassword')
                self.m_staticPasswordexpiration = wx.FindWindowByName('m_staticPasswordexpiration')
                self.m_datePickerpasswordexp = wx.FindWindowByName('m_datePickerpasswordexp')
                self.m_checkBoxactivated = wx.FindWindowByName('m_checkBoxactivated')
                self.m_checkBoxgetrevision = wx.FindWindowByName('m_checkBoxgetrevision')
                self.m_checkBoxgetpopsheet = wx.FindWindowByName('m_checkBoxgetpopsheet')
                self.m_checkBoxengineer = wx.FindWindowByName('m_checkBoxengineer')
                self.m_checkBoxcadeng = wx.FindWindowByName('m_checkBoxcadeng')
                self.m_checkBoxprojecteng = wx.FindWindowByName('m_checkBoxprojecteng')
                self.m_checkBoxprojectlead = wx.FindWindowByName('m_checkBoxprojectlead')
                self.m_checkBoxscheduler = wx.FindWindowByName('m_checkBoxscheduler')
                self.m_checkBoxapprovefirst = wx.FindWindowByName('m_checkBoxapprovefirst')
                self.m_checkBoxApprovesecond = wx.FindWindowByName('m_checkBoxApprovesecond')
                self.m_checkBoxReadonly = wx.FindWindowByName('m_checkBoxReadonly')
                self.m_checkBoxAdmin = wx.FindWindowByName('m_checkBoxAdmin')
                self.m_buttonSubmit = wx.FindWindowByName('m_buttonSubmit')
                self.m_export_to_excel=wx.FindWindowByName('m_export_to_excel')
                self.fetchDB()
                adminuser=[]
                selected_user=self.name.GetStringSelection()
                self.dbCursor.execute('select name from employees where Admin=1')
                DB_values=self.dbCursor.fetchone()
                while DB_values!= None:
                        adminuser.append(DB_values[0])
                        DB_values=self.dbCursor.fetchone()
                
                if selected_user not in adminuser:                         
                         self.createuserpanel.Show(False)                         
                         self.m_staticName.Hide()
                         self.m_textName.Hide()
                         self.m_staticdepartment.Hide()
                         self.m_combodepartment.Hide()
                         self.m_staticEmail.Hide()
                         self.m_textEmail.Hide()
                         self.m_staticPassword.Hide()
                         self.m_textpassword.Hide()
                         self.m_staticPasswordexpiration.Hide()
                         self.m_datePickerpasswordexp.Hide()
                         self.m_checkBoxactivated.Hide()
                         self.m_checkBoxgetrevision.Hide()
                         self.m_checkBoxgetpopsheet.Hide()
                         self.m_checkBoxengineer.Hide()
                         self.m_checkBoxcadeng.Hide()
                         self.m_checkBoxprojecteng.Hide()
                         self.m_checkBoxprojectlead.Hide()
                         self.m_checkBoxscheduler.Hide()
                         self.m_checkBoxapprovefirst.Hide()
                         self.m_checkBoxApprovesecond.Hide()
                         self.m_checkBoxReadonly.Hide()
                         self.m_checkBoxAdmin.Hide()
                         self.m_buttonSubmit.Hide()
                         self.m_export_to_excel.Hide()
                         
                           
                else:
                        
                        self.createuserpanel=wx.FindWindowByName('createuserpanel')
                        self.Bind(wx.EVT_BUTTON, self.insertdetails, id=xrc.XRCID('m_buttonSubmit'))
                        self.Bind(wx.EVT_BUTTON, self.OnBtnExportData, id=xrc.XRCID('m_export_to_excel'))
                        self.m_textName = wx.FindWindowByName('m_textName')                
                        self.m_combodepartment = wx.FindWindowByName('m_combodepartment')
                        self.m_textEmail = wx.FindWindowByName('m_textEmail')
                        self.m_textpassword= wx.FindWindowByName('m_textpassword')
                        self.m_datePickerpasswordexp=wx.FindWindowByName('m_datePickerpasswordexp')
                        self.m_checkBoxactivated=wx.FindWindowByName('m_checkBoxactivated')
                        self.m_checkBoxgetrevision=wx.FindWindowByName('m_checkBoxgetrevision')
                        self.m_checkBoxgetpopsheet=wx.FindWindowByName('m_checkBoxgetpopsheet')
                        self.m_checkBoxengineer=wx.FindWindowByName('m_checkBoxengineer')
                        self.m_checkBoxcadeng=wx.FindWindowByName('m_checkBoxcadeng')
                        self.m_checkBoxprojecteng=wx.FindWindowByName('m_checkBoxprojecteng')
                        self.m_checkBoxprojectlead=wx.FindWindowByName('m_checkBoxprojectlead')
                        self.m_checkBoxscheduler=wx.FindWindowByName('m_checkBoxscheduler')
                        self.m_checkBoxapprovefirst=wx.FindWindowByName('m_checkBoxapprovefirst')
                        self.m_checkBoxApprovesecond=wx.FindWindowByName('m_checkBoxApprovesecond')
                        self.m_checkBoxReadonly=wx.FindWindowByName('m_checkBoxReadonly')
                        self.m_checkBoxAdmin = wx.FindWindowByName('m_checkBoxAdmin')
                        self.m_buttonSubmit=wx.FindWindowByName('m_buttonSubmit')
                        self.m_export_to_excel=wx.FindWindowByName('m_export_to_excel')
                        
                        #Database connection variables
                        self.AENames = []
                        #self.dbCursor = None               
                        
        def fetchDB(self):

                #Database connection variables
                conn = db.connect_to_eng04_database() #create a new database connection for admin.                
                self.dbCursor = conn.cursor()
                self.db_dropdown()
               
                
        

        #method for binding dropdown with departments
        def db_dropdown(self):
                
              del self.AENames[:]                           
              self.dbCursor.execute('select department from departments')
              row = self.dbCursor.fetchone()
              
              while row != None:                
                      self.AENames.append(row[0])
                      row = self.dbCursor.fetchone()
                      
              self.m_combodepartment.SetItems(self.AENames)             


        #method for insert or update record to db
        def insertdetails(self,event):        
           activated=self.m_checkBoxactivated.GetValue()
           name=str(self.m_textName.GetValue())
           password= str(self.m_textpassword.GetValue())           
           passwordexp=self.m_datePickerpasswordexp.GetValue()                      
           dt=parse(str(passwordexp))           
           email=str(self.m_textEmail.GetValue())           
           department=str(self.m_combodepartment.GetValue())           
           revision_notice=self.m_checkBoxgetrevision.GetValue()           
           popsheet=self.m_checkBoxgetpopsheet.GetValue()           
           isenginner=self.m_checkBoxengineer.GetValue()           
           cad=self.m_checkBoxcadeng.GetValue()           
           projectenginner=self.m_checkBoxprojecteng.GetValue()
           projectlead=self.m_checkBoxprojectlead.GetValue()
           scheduler=self.m_checkBoxscheduler.GetValue()         
           approvefirst=self.m_checkBoxapprovefirst.GetValue()           
           approvesecond=self.m_checkBoxApprovesecond.GetValue()          
           readonly= self.m_checkBoxReadonly.GetValue()
           admin=self.m_checkBoxAdmin.GetValue()
           self.dbCursor.execute("select max(id) from  [employees]")           
           ID=self.dbCursor.fetchone()                      
           u_ID=int(ID[0])                    
           id=u_ID+1           
           self.dbCursor.execute("select * from  [employees] where email = ?",(email,))                    
           listvalue=self.dbCursor.fetchone()           
           if listvalue:
                         if not name or not department or not email or not password:
                               msgbox = wx.MessageBox('Please enter your details', 'Alert')
                         else:
                                 if passwordexp.IsValid() == True and passwordexp.IsLaterThan(wx.DateTime_Today()) == False:      
                                           msgbox = wx.MessageBox('The password expiration date cannot be in past!', 'Alert')
                                 else:
                                         
                                       if ('@LENNOXINTL.COM' in email.upper()) or ('@HEATCRAFTRPD.COM' in email.upper()):
                                               regex=re.match('^[_a-zA-Z0-9-]+(\.[_a-zA-Z0-9-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z]{2,3})$',email.upper())
                                               if regex:
                                                        self.dbCursor.execute("""UPDATE employees SET activated=?,name=?,password=?,password_expiration=?,
                                                       email=?,department=?,
                                                       gets_revision_notice=?
                                                      ,gets_pop_sheet_notice=?
                                                      ,is_engineer=?
                                                      ,is_cad_designer=?
                                                      ,is_project_engineer=?
                                                      ,is_project_lead=?
                                                      ,is_scheduler=?
                                                      ,can_approve_first=?
                                                      ,can_approve_second=?
                                                      ,is_readonly=?,
                                                      Admin=?
                                                      WHERE email=?""",
                                                      (activated,name, password,dt,email,department,revision_notice,popsheet,isenginner,cad,projectenginner,projectlead,scheduler,approvefirst,approvesecond,readonly,admin,email))
                                                        self.dbCursor.commit() 
                                                        msgbox = wx.MessageBox('Data successfully updated', 'Alert')
                                               else:
                                                        msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com', 'Alert')
                                                        
                                       else:
                                               msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com', 'Alert')


           else:
                   if  passwordexp.IsValid() == True and passwordexp.IsLaterThan(wx.DateTime_Today()) == False:      
                                           msgbox = wx.MessageBox('The password expiration date cannot be in past!', 'Alert')
                                    
                   else:
                           if ('@LENNOXINTL.COM' in email.upper()) or ('HEATCRAFTRPD.COM' in email.upper()):                                   
                                   regex=re.match('^[_a-zA-Z0-9-]+(\.[_a-zA-Z0-9-]+)*@[a-zA-Z0-9-]+(\.[a-zA-Z]{2,3})$',email.upper())
                                   if regex:
                                           
                                               self.dbCursor.execute("""insert into employees(activated,name,password,password_expiration,email,department,gets_revision_notice
                                              ,[gets_pop_sheet_notice]
                                              ,[is_engineer]
                                              ,[is_cad_designer]
                                              ,[is_project_engineer]
                                              ,[is_project_lead]
                                              ,[is_scheduler]
                                              ,[can_approve_first]
                                              ,[can_approve_second]
                                              ,[is_readonly],[Admin]) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                             [activated,name, password,dt,email,department,revision_notice,popsheet,isenginner,cad,projectenginner,projectlead,scheduler,approvefirst,approvesecond,readonly,admin])                                                  
                                               self.dbCursor.commit()
                                               msgbox = wx.MessageBox('Data successfully inserted', 'Alert')
                                   else:
                                        msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com', 'Alert')       
                           else:
                                   msgbox = wx.MessageBox('Email should be like example@Lennoxintl.com or example@Heatcraftrpd.com', 'Alert')

                        
        #method for export the excel of employee details                     
        def OnBtnExportData(self,event=None):
                excel = Dispatch('Excel.Application')
                excel.Visible = True
                wb = excel.Workbooks.Add()                
                self.dbCursor.execute("""
                SELECT Admin,activated AS Activated,
                name AS [Name],
                password AS[Password],
                password_expiration AS [Password Expiration],
                email AS [Email],
                department AS [Department],gets_revision_notice AS [Gets Revision Notice],gets_pop_sheet_notice as [Get Pop Sheet Notice],is_engineer as [Engineer],
                is_cad_designer as [CAD],is_project_engineer as [Project Engineer],is_project_lead as [Project Lead],is_scheduler as [Scheduler],can_approve_first as [Approve First],
                can_approve_second as [Approve Second],is_readonly as [Readonly] FROM  [employees] """)
                column=[ i[0] for i in self.dbCursor .description ]
                rows=self.dbCursor .fetchall()
                print len(column)
                print len(rows)
                    
                wb.ActiveSheet.Cells(1, 1).Value = rows
                wb.ActiveSheet.Columns(1).AutoFit()
                R=len(rows)
                C=len(column)
                 #Write the header
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(1,C))
                excel_range.Value = column
                
                #specify the excel range that the main data covers
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(2, 1), wb.ActiveSheet.Cells(R+1,C))
                excel_range.Value = rows

                #Autofit the columns
                excel_range = wb.ActiveSheet.Range(wb.ActiveSheet.Cells(1, 1),wb.ActiveSheet.Cells(R+1,C))
                excel_range.Columns.AutoFit()
                
               
##                with open(save_path, 'wb') as csvfile:                                                                
##                        writer = csv.writer(csvfile)                        
##                        writer.writerow([ i[0] for i in self.dbCursor.description ]) # heading row                        
##                        writer.writerows(self.dbCursor.fetchall())
##                        
##                msgbox = wx.MessageBox('Export successfully', 'Alert')                

