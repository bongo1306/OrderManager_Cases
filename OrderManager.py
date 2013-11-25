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

import ConfigParser #for reading local config data (*.cfg)

import datetime as dt

import BetterListCtrl
import General as gn
import Database as db
