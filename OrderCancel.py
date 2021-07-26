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
 
user = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
orderRefKey = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
errCode = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_VARIANT, 0)
 
user.value ="r151681"
orderRefKey.value="gS046979202"

q.CancelByKey(user, orderRefKey , errCode)
print("user" + str(user.value) + ", orderRefKey="+orderRefKey.value + ", errCode=" + " " + str(errCode.value))

