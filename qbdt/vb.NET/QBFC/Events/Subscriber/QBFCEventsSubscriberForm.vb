Option Strict Off
Option Explicit On
Friend Class QBFCEventsSubscriber
	Inherits System.Windows.Forms.Form
	Private Sub Subscribe_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Subscribe.Click
		
		DoSubscribeEvents()
		
	End Sub
	
	Private Sub Unsubscribe_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Unsubscribe.Click
		
		DoUnSubscribeEvents()
		
	End Sub
End Class