#!/usr/bin/env python
# -*- coding: utf-8 -*- 

__author__ = "aldria@post.cz"
__date__ = "$19.11.2015$"

# Definice guid èíselníkù z GxDoc.chm
cRollStoreCardCategories = 'K40Q4IS15VDL342P01C0CX3FCC' # Èíselník typù skladových karet 
cRollVATRates = 'KE4KIBA3Y3CL33N2010DELDFKK' # Èíselník sazeb DPH

import win32com.client

abraOle = win32com.client.Dispatch('AbraOLE.Application')
abraOle.login('supervisor','')

def getIDFromRoll(ole ,rollGUID, text, field):
    mRoll = ole.GetRoll(rollGUID, 0)
    state, mID = mRoll.Find(field, text, '')
    if state:
        return mID

storeCardObject = abraOle.CreateObject("@StoreCard")
storeCardData = abraOle.CreateValues("@StoreCard")

storeCardObject.PrefillValues(storeCardData)
storeCardData.SetValueByName("Code", "001")
storeCardData.SetValueByName("Name", "hendrerit ante")
storeCardData.SetValueByName("StoreCardCategory_ID", getIDFromRoll(abraOle, cRollStoreCardCategories, '01', 'Code')) 
storeCardData.SetValueByName("VatRate_ID", getIDFromRoll(abraOle, cRollVATRates, '21', 'Tariff')) 

storeCardID = storeCardObject.CreateNewFromValues(storeCardData)

del abraOle 

print 'store card id is:  ' + storeCardID