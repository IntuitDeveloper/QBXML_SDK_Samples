Imports System.Linq

Module QBFCEventsCallbackMain
    Public frm As New QBFCEventsCallbackForm
    Sub Main()
        Dim strArg() As String

        strArg = Environment.GetCommandLineArgs()

        If (strArg.Count = 2) Then
            If (strArg(1) = "-Embedding") Then
                'called from QBDT so create the form here
                frm.Show()
            End If
        End If

        QBFCEventsCallbackCOMServer.Instance.Run()
    End Sub

End Module
