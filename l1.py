import win32com.client
import time
from win32com.client import pythoncom, VARIANT
from time import gmtime, strftime
 
# Equity Order Entry Example
q = win32com.client.Dispatch("REDI.Query")
 
# Prepare a variable which can handle returned values.
msg1 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
symbolVar = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
fieldVar = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
tgtVarName = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
 
vTable = "L1"
vWhere = "true"
tmpVal = q.Submit(vTable, vWhere, msg1)
 
symbolVar.value = "IBM"
ret = q.GetL1Value(symbolVar, "Last", tgtVarName)
print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + " Symbol=" + tgtVarName.value + " Last" + "=" + str(symbolVar.value) + " success=" + str(ret))

symbolVar.value = "T"
ret =  q.GetL1Value(symbolVar, "Ask", tgtVarName)
print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + " Symbol=" + tgtVarName.value + " " + "Ask" + "=" + str(symbolVar.value) + " success=" + str(ret))

symbolVar.value = "BA"
ret = q.GetL1Value(symbolVar, "Bid", tgtVarName)
print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + " Symbol=" + tgtVarName.value + " " + "Bid" + "=" +  str(symbolVar.value) + " success=" + str(ret))

symbolVar.value = "InvalidSymbol"
ret = q.GetL1Value(symbolVar, "Bid", tgtVarName)
print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + " Symbol=" + tgtVarName.value + " " + "InvalidSymbol" + "=" +  str(symbolVar.value) + " success=" + str(ret))