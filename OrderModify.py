# |-----------------------------------------------------------------------------
# |            This source code is provided under the Apache 2.0 license      --
# |  and is provided AS IS with no warranty or guarantee of fit for purpose.  --
# |                See the project's LICENSE.md for details.                  --
# |           Copyright (C) 2021 Refinitiv. All rights reserved.              --
# |-----------------------------------------------------------------------------
"""
# Load COM interface
import win32com.client

# Equity Order Entry Example
o = win32com.client.Dispatch("REDI.ORDER")

o.SetOrderKey("r151681", "gS042857202")
o.ClientData = "GyGy"

# Prepare a variable which can handle returned values from submit method of the order object.
msg = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_VARIANT, None)

# Submit modification
result = o.Submit(msg)

print(result) # 'True' if order submission was successful; otherwise 'False'
print(msg)    # message from sumbit 
