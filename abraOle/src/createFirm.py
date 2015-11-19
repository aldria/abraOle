#!/usr/bin/env python
# -*- coding: utf-8 -*- 

__author__ = "aldria@post.cz"
__date__ = "$15.11.2015 16:46:22$"

import win32com.client

abraOle = win32com.client.Dispatch('AbraOLE.Application')
abraOle.login('supervisor','')

firmObject = abraOle.CreateObject("@Firm")
firmData = abraOle.CreateValues("@Firm")

firmObject.PrefillValues(firmData)
firmData.SetValueByName("Name", "Acme Corporation")
firmData.SetValueByName("OrgIdentNumber", "123456")
firmData.SetValueByName("VatIdentNumber", "CC123456")

addressData = firmData.GetValueByName("ResidenceAddress_ID")
addressData.SetValueByName("Street", "Street")

firmID = firmObject.CreateNewFromValues(firmData)

del abraOle 

print 'firm id is:  ' + firmID