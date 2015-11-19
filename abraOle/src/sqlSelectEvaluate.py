#!/usr/bin/env python
# -*- coding: cp1250 -*- 

__author__ = "aldria@post.cz"
__date__ = "$19.11.2015$"

import win32com.client

abraOle = win32com.client.Dispatch('AbraOLE.Application')
abraOle.login('supervisor','')
firmObject = abraOle.CreateObject("@Firm")
mStrings = abraOle.CreateStrings()
SQLResult = abraOle.SQLSelect("select id from firms", mStrings)
print mStrings.Count()
for x in range(0, mStrings.Count()):
    print mStrings.Strings(x)
    print (firmObject.Evaluate(mStrings.Strings(x), 'name')).encode('cp1250')

del abraOle 