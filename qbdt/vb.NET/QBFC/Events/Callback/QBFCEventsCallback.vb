Option Strict Off
Option Explicit On
Friend Class QBFCEventsCallbackForm
	Inherits System.Windows.Forms.Form
	Private Sub Query_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Query.Click
		GetAllCustomers()
	End Sub
End Class