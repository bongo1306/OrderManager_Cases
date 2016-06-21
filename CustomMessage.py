import wx
import os
from wx import xrc
ctrl = xrc.XRCCTRL


class MessageDialog(wx.Dialog):
	def __init__(self, parent):
		#load XRC description
		pre = wx.PreDialog()
		res = xrc.XmlResource.Get()
		res.LoadOnDialog(pre, parent, "MessageDialog")
		self.PostCreate(pre)
		ImgPath = str(os.path.split(__file__)[0]) + "\Images\keys.png"
		self.SetIcon(wx.Icon(ImgPath, wx.BITMAP_TYPE_PNG))

		#bindings
		self.Bind(wx.EVT_BUTTON, self.OnBtnOK, id=xrc.XRCID('m_BtnOK'))
		self.Bind(wx.EVT_CLOSE, self.OnClose)

		#Get the windows
		self.m_MsgTextCtrl = wx.FindWindowByName('m_MsgTextCtrl')
		self.m_MsgTextCtrl.SetFocus()
		self.m_MsgTextCtrl.ShowNativeCaret(False)

	def OnBtnOK(self, event):
		self.Destroy()

	def OnClose(self, event):
		self.Destroy()