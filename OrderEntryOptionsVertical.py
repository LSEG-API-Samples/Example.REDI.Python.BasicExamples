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

# Load COM interface
import win32com.client
from win32com.client import pythoncom, VARIANT

# Option Order Entry Example
o = win32com.client.Dispatch("REDI.COMPLEXORDER")

# Complex options order header
o.Strategy  = "Vertical"
o.SetSymbol(0,"SPX")
#o.SetRootSymbol(0,"SPXW")
o.SetExchange(0,"DEM2 DMA")
o.SetQuantity(0, 1)
o.SetPriceType(0,"Market")
o.SetTIF(0,"Day")
o.SetAccount(0,"EQUITY-TR")

# Leg 1 of vertical spread
o.SetSide(1,"Buy")
o.SetPosition(1,"Open")
o.SetOptType(1,"Call")
o.SetMonth(1,"Aug 13 '21")
o.SetStrike(1,4405.00)

# Leg 2 of vertical spread
o.SetSide(2,"Sell")
o.SetPosition(2,"Open")
o.SetOptType(2,"Call")
o.SetMonth(2,"Aug 13 '21")
o.SetStrike(2,4400.00)

# Prepare a variable which can handle returned values from submit method of the order object.
msg = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_VARIANT, None)

# Send an options order
result = o.Submit(msg)

print(result) # 'True' if order submission was successful; otherwise 'False'
print(msg)    # message from sumbit 
