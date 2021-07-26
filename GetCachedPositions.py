# |-----------------------------------------------------------------------------
# |            This source code is provided under the Apache 2.0 license      --
# |  and is provided AS IS with no warranty or guarantee of fit for purpose.  --
# |                See the project's LICENSE.md for details.                  --
# |           Copyright (C) 2021 Refinitiv. All rights reserved.              --
# |-----------------------------------------------------------------------------

# |-----------------------------------------------------------------------------
# | Please be informed, that this example uses python library win32com 		  --
# | that is not provided or supported by Refinitiv							  --
# |-----------------------------------------------------------------------------

import win32com.client
import time
from win32com.client import pythoncom, VARIANT
from time import gmtime, strftime
 
q = win32com.client.Dispatch("REDI.Query")
 
# Prepare a variable which can handle returned values from submit method of the order object.
row = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
column = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
cellValue = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
errCode = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)

#add watch on accounts 
myaccounts = ["EQUITY-TR"] 
for account in myaccounts:
	tmpVal = q.AddWatch("2", "", account, errCode);
	print("account=" + account +  ", errCode=" + " " + str(errCode.value))

# Prepare the query
vTable = "Position"
vWhere = "true"
tmpVal = q.Submit(vTable, vWhere, errCode)

# Find out the number of available rows
rowCount = q.RowCount
print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + ": "+str(rowCount)+" rows present")
 
for i in range(0, rowCount):

	row.value = i
	column.value = "Account"
	q.GetCell(row, column, cellValue , errCode)
	print("row=" + str(row.value) + ", column="+column.value + ", cellValue=" + cellValue.value + ", errCode=" + " " + str(errCode.value))

	column.value = "Symbol"
	q.GetCell(row, column, cellValue , errCode)
	print("row=" + str(row.value) + ", column="+column.value + ", cellValue=" + str(cellValue.value) + ", errCode=" + " " + str(errCode.value))

	column.value = "Position"
	q.GetCell(row, column, cellValue , errCode)
	print("row=" + str(row.value) + ", column="+column.value + ", cellValue=" + str(cellValue.value) + ", errCode=" + " " + str(errCode.value))

	column.value = "Value"
	q.GetCell(row, column, cellValue , errCode)
	print("row=" + str(row.value) + ", column="+column.value + ", cellValue=" + str(cellValue.value) + ", errCode=" + " " + str(errCode.value))

	print("===>");
