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
import codecs
import collections
import datetime
import itertools
import StringIO

import xlwt

import wx
import MainGui
import AboutDialog

import arial10

VERSTR = 'ToCParser 1.01'

# Implementing MainFrame
class MainFrame(MainGui.MainFrame):
    def __init__(self, parent):
	MainGui.MainFrame.__init__(self, parent)
        self.SetTitle(VERSTR)
        self.config = wx.Config('ToCParser', 'ToCParser');
        self.config.SetRecordDefaults(True)
        remember = self.config.ReadBool('RememberFilename', True)
        self.m_checkBoxRememberFilename.SetValue(remember)
        unnumbered = self.config.ReadBool('ConvertUnnumbered', False)
        self.m_checkBoxConvertUnnumbered.SetValue(unnumbered)

        if remember:
            lastfile = self.config.Read('LastFilename', '')
            if lastfile:
                self.m_filePickerDest.SetPath(lastfile)

    def OnCheckBoxConvertUnnumbered(self, event):
        event.Skip()
        unnumbered = self.m_checkBoxConvertUnnumbered.GetValue()
        self.config.WriteBool('ConvertUnnumbered', unnumbered)

    def OnTextToC(self, event):
        event.Skip()
        tocfile = StringIO.StringIO(self.m_textCtrlToC.GetValue())
        nlines = 0
        for line in tocfile:
            nlines = nlines + 1

        self.m_statusBar.SetStatusText('Pasted ToC with %d lines' % nlines)

    def OnButtonClickAbout(self, event):
        event.Skip()
	dlg = AboutDialog.AboutDialog(self)
	ret = dlg.ShowModal()

    def OnCheckBoxRememberFilename(self, event):
        event.Skip()
        remember = self.m_checkBoxRememberFilename.GetValue()
        self.config.WriteBool('RememberFilename', remember)

    def OnFileChangedDestFile(self, event):
        event.Skip()
        if self.m_checkBoxRememberFilename.GetValue():
            destfile = self.m_filePickerDest.GetPath()
            self.config.Write('LastFilename', destfile)

    def OnButtonClickClearRegistryAndExit(self, event):
        event.Skip()
        self.config.DeleteAll()
        self.Close()

    def LogClear(self):
        self.m_textCtrlConvLog.Clear()

    def LogMsg(self, msg):
        self.m_textCtrlConvLog.AppendText('%s\r\n' %msg)

    def OnButtonClickConvertToExcel(self, event):
        headingxf = xlwt.easyxf('font: bold on; align: wrap on, vert centre, horiz center')
        generalxf = xlwt.easyxf()
        numxf = xlwt.easyxf(num_format_str='0')
        textxf = xlwt.easyxf(num_format_str='@')

        coldata = collections.defaultdict(list)

        event.Skip()
        self.LogClear()

        dstname = self.m_filePickerDest.GetPath()
        if not dstname:
            self.LogMsg('Conversion failed: No destination file selected')
            return

        toctxt = self.m_textCtrlToC.GetValue()
        tocfile = StringIO.StringIO(toctxt)

        dst = xlwt.Workbook()
        ws = dst.add_sheet('Requirements')
        # Write header
        itemid = 0
        colidx = itertools.count()
        colcol = itertools.count()
        ws.write(itemid, next(colidx), 'Id', headingxf)
        coldata[next(colcol)].append('Id')
        ws.write(itemid, next(colidx), 'Section', headingxf)
        coldata[next(colcol)].append('Section')
        ws.write(itemid, next(colidx), 'Description', headingxf)
        coldata[next(colcol)].append('Description')
        ws.write(itemid, next(colidx), 'Page', headingxf)
        coldata[next(colcol)].append('Page')
        ws.write(itemid, next(colidx), 'Fully Compliant', headingxf)
        coldata[next(colcol)].append('Compliant')
        ws.write(itemid, next(colidx), 'Partially Compliant', headingxf)
        coldata[next(colcol)].append('Partially')
        ws.write(itemid, next(colidx), 'Not Compliant', headingxf)
        coldata[next(colcol)].append('Compliant')
        ws.write(itemid, next(colidx), 'Expected Timeframe', headingxf)
        coldata[next(colcol)].append('Timeframe')
        ws.write(itemid, next(colidx), 'Comment', headingxf)
        coldata[next(colcol)].append('Comment000000000000000000000000000000000000000000000')

        totalcols = next(colidx) - 1

        try:
            linenum = 0
            itemid = itemid + 1
            unnumbered = self.m_checkBoxConvertUnnumbered.GetValue()
            self.LogMsg('Conversion started: %s' % datetime.datetime.now().isoformat())
            for line in tocfile:
                linenum = linenum + 1
                if not linenum % 10:
                    self.LogMsg('Processed lines: %d-%d' % (linenum - 10 + 1, linenum))
                line = line.rstrip('\r\n')
                if not line:
                    self.LogMsg('Skipped empty line %d' % linenum)
                    continue
                splitted = line.split()
                secnum = splitted[0].rstrip()
                # if not secnum[0].isdigit():
                if not secnum[-1] == '.':
                    if not unnumbered:
                        self.LogMsg('Skipped unnumbered section %d: %s' % (linenum, line))
                        continue
                    secnum = ''
                else:
                    secnum = secnum.rstrip('.').lstrip()
                secpage = splitted[-1]
                sectitle = ' '.join(splitted[1:-1])

                colidx = itertools.count()
                colcol = itertools.count()
                ws.write(itemid, next(colidx), label=itemid, style=numxf)
                coldata[next(colcol)] = str(itemid)
                ws.write(itemid, next(colidx), label=secnum, style=textxf)
                coldata[next(colcol)] = str(secnum)
                ws.write(itemid, next(colidx), label=sectitle, style=textxf)
                coldata[next(colcol)].append(sectitle)
                ws.write(itemid, next(colidx), label=int(secpage), style=numxf)
                coldata[next(colcol)].append(str(secpage))
                itemid = itemid + 1

            rem = linenum % 10
            if rem:
                self.LogMsg('Processed lines: %d-%d' % (linenum - rem, linenum))

        except Exception, e:
            self.LogMsg('Aborting - Error during conversion: %s' % str(e))
        finally:
            try:

                for col in xrange(totalcols + 1):
                    maxw = 0
                    for data in coldata[col]:
                        w = arial10.fitwidth(data)
                        if w > maxw:
                            maxw = w

                    # Fit the column
                    ws.col(col).width = maxw + 1000

                ws.set_panes_frozen(True) # frozen headings instead of split panes
                ws.set_horz_split_pos(1) # in general, freeze after last heading row
                ws.set_remove_splits(True) # if user does unfreeze, don't leave a split there
                dst.save(dstname)
            except Exception, e:
                self.LogMsg('Error saving excel sheet: %s' % str(e))

        tocfile.close()
        self.LogMsg('Conversion ended: %s' % datetime.datetime.now().isoformat())
