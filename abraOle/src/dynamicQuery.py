#!/usr/bin/env python
# -*- coding: utf-8 -*- 

__author__ = "aldria@post.cz"
__date__ = "$19.11.2015$"

# Definice guid dynamického dotazu z GxDoc.chm
cDynSQLFirms = 'W0DR1FTE3JD13ACL03KIU0CLP4' # dyn sql firms 

import win32com.client

abraOle = win32com.client.Dispatch('AbraOLE.Application')
abraOle.login('supervisor','')
firmObject = abraOle.CreateObject("@Firm")
mStrings = abraOle.CreateStrings()

mDynSQL = abraOle.CreateCustomCommand(cDynSQLFirms) 
mCond = mDynSQL.ConstraintByID("ID")
mCond.UsedKind = 1
mCond.Value = "'" + "AAA1000000" + "'"
mDataset = mDynSQL.RowsetByName("MAIN")
mDataset.UsedFields = "Name" + '\r\n' + "Code" 
mDataset.Used = True
mDynSQL.Execute()

while not mDataset.EOF:
  print (mDataset.Data.ValueByName("Name")).encode('cp1250')
  print (mDataset.Data.ValueByName("Code")).encode('cp1250')
  mDataset.Next()

del abraOle 