Option Strict Off
Option Explicit On
Imports System.Linq

Module MainMod
    '
    ' Since we are an ActiveX COM server, if we get started interactively we need some
    ' where to begin, that's here and we just bring up our form.
    '
    'UPGRADE_WARNING: Application will terminate when Sub Main() finishes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E08DDC71-66BA-424F-A612-80AF11498FF8"'
    Public frm As New MenuEventSample

    Public Sub Main()
        frm.Show()
        QBFCMenuEventsCallbackCOMServer.Instance.Run()
    End Sub
End Module