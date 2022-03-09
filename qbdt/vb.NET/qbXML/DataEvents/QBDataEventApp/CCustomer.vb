Option Strict Off
Option Explicit On
Friend Class CCustomer
	'
	' this is just a trivial little object (note all variables are public so we don't even
	' have to bother with Property Let/Get methods) to hold the information we care about
	' in a customer -- kind of simulates a database record.
	'
	
	Public name As String
	Public Phone As String
	Public Email As String
	Public Balance As String
	Public index As String
	Public ListID As String
	
	Sub Init()
		name = ""
		Phone = ""
		Email = ""
		Balance = "0.00"
		index = ""
		ListID = ""
	End Sub
End Class