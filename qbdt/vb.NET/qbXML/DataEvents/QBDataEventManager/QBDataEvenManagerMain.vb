Module QBDataEvenManagerMain
    Public frm As New QBDataEventManagerDisplay
    Sub Main(args As String())
        Dim strArg() As String

        strArg = Environment.GetCommandLineArgs()

        If (strArg.Count = 2) Then
            If (strArg(1) = "-Embedding") Then
                'called from QBDT so create the form here
                frm.Show()
            End If
        End If

        QBDataEventManagerCOMServer.Instance.Run()
    End Sub

End Module
