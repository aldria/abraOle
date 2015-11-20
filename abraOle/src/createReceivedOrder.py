#!/usr/bin/env python
# -*- coding: utf-8 -*- 

__author__ = "aldria@post.cz"
__date__ = "$15.11.2015 16:46:22$"

cDocQueueID = '1000000101' #ID �ady doklad� objedn�vky p�ijat�
cFirmID = 'AAA1000000' #ID firmy "#Bez p��slu�nosti k firm�"
cDivisionID = '1000000101' #ID st�ediska

import win32com.client
from datetime import datetime

abraOle = win32com.client.Dispatch('AbraOLE.Application')
abraOle.login('supervisor','')

receivedOrderObject = abraOle.CreateObject('@ReceivedOrder')
receivedOrderData = abraOle.CreateValues('@ReceivedOrder')
receivedOrderObject.PrefillValues(receivedOrderData)
receivedOrderData.SetValueByName('Description', 'Testovac� objedn�vka')
receivedOrderData.SetValueByName('DocQueue_ID', cDocQueueID)
receivedOrderData.SetValueByName('Firm_ID', cFirmID)           
receivedOrderData.SetValueByName('DocDate$DATE', datetime(2015, 11, 20, 0,0,0).toordinal() - datetime(1899, 12, 30, 0, 0, 0).toordinal())           
receivedOrderData.SetValueByName('IsRowDiscount', True)
receivedOrderData.SetValueByName('PricesWithVAT', False)
            
receivedOrderRowDataCollection = receivedOrderData.ValueByName('Rows')
receivedOrderRowData = abraOle.CreateValues('@ReceivedOrderRow');
receivedOrderRowData.SetValueByName('RowType', 0)
receivedOrderRowData.SetValueByName('Division_ID', cDivisionID)
receivedOrderRowData.SetValueByName('Text', 'Testovac� objedn�vka')
receivedOrderRowDataCollection.Add(receivedOrderRowData) 

receivedOrderID = receivedOrderObject.CreateNewFromValues(receivedOrderData)       
           
del abraOle 

print 'received order ID:  ' + receivedOrderID