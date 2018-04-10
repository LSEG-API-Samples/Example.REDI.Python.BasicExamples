import win32com.client
import time
from win32com.client import pythoncom, VARIANT
from time import gmtime, strftime
 
# Equity Order Entry Example
q = win32com.client.Dispatch("REDI.Query")
 
# Prepare a variable which can handle returned values from submit method of the order object.
msg1 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
msg2 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
msg3 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
msg4 = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
 
vTable = "L1"
vWhere = "true"
tmpVal = q.Submit(vTable, vWhere, msg1)
 
msg2.value = "IBM"
msg3.value = "Last"
q.GetL1Value(msg2, msg3, msg4)
print(strftime("%Y-%m-%d %H:%M:%S", gmtime()) + " " + msg3.value + " " + msg4.value + " " + msg2.value)
