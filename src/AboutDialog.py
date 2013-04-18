#!/usr/bin/env python
# -*- coding: latin-1; py-indent-offset:4 -*-
################################################################################
# 
# This file is part of TOCParser
#
# TOCParser is an interface to automate parsing a Table of Contents and
# export it to CSV (or Microsoft (R) Excel)
#
# Copyright (C) 2012 Daniel Rodriguez
#
# TOCParser is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# TOCParser is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with TOCParser. If not, see <http://www.gnu.org/licenses/>.
#
################################################################################
import wx
import MainGui

# Implementing AboutDialog
class AboutDialog(MainGui.AboutDialog):
    def __init__(self, parent):
	MainGui.AboutDialog.__init__(self, parent)
	
    # Handlers for AboutDialog events.
    def OnButtonClickOk(self, event):
	self.EndModal(wx.ID_OK)
