Option Strict Off
Option Explicit On
Module MainMod
	'
	' Since we are an ActiveX COM server, if we get started interactively we need some
	' where to begin, that's here and we just bring up our form.
	'
	
	Public Sub Main()
        Dim Form As New MenuEventSample
        Form.ShowDialog()
    End Sub
End Module