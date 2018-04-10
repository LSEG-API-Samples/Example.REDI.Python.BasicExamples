# Load COM interface
import win32com.client

# Equity Order Entry Example
o = win32com.client.Dispatch("REDI.ORDER")
o.Side      = "Buy"
o.symbol    = "SBUX"
o.Exchange  = "BETA DMA"
o.Quantity  = "1"
o.PriceType = "Limit"
o.Price     = "0.01"
o.TIF       = "Day"
o.Account   = "00999900"
o.Ticket    = "Bypass"

# Prepare a variable which can handle returned values from submit method of the order object.
msg = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_VARIANT, None)

# Send an options order
result = o.Submit(msg)

print result # 'True' if order submission was successful; otherwise 'False'
print msg    # message from sumbit 
