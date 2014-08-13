Attribute VB_Name = "MainMod"
'
' Since we are an ActiveX COM server, if we get started interactively we need some
' where to begin, that's here and we just bring up our form.
'
Public Sub Main()
    Load MenuEventSample
    MenuEventSample.Show
End Sub

