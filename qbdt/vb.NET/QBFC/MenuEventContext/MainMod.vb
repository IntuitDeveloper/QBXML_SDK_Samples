Option Strict Off
Option Explicit On
Imports System.Linq

Module MainMod
    '
    ' Since we are an ActiveX COM server, if we get started interactively we need some
    ' where to begin, that's here and we just bring up our form.
    '
    
    Public frm As New MenuEventSample

    Public Sub Main()
        frm.Show()
        QBFCMenuEventsCallbackCOMServer.Instance.Run()
    End Sub
End Module