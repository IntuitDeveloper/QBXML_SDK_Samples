

Imports System.Runtime.InteropServices

Friend Class NativeHelper
    <DllImport("user32.dll")>
    Friend Shared Function GetMessage(<Out()> ByRef lpMsg As MSG,
                                      ByVal hWnd As IntPtr,
                                      ByVal wMsgFilterMin As UInt32,
                                      ByVal wMsgFilterMax As UInt32) As Boolean
    End Function



    <DllImport("user32.dll")>
    Friend Shared Function TranslateMessage(<[In]()> ByRef lpMsg As MSG) As Boolean
    End Function


    <DllImport("user32.dll")>
    Friend Shared Function DispatchMessage(<[In]()> ByRef lpMsg As MSG) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Friend Shared Function PostThreadMessage(ByVal idThread As UInt32,
                                             ByVal Msg As UInt32,
                                             ByVal wParam As UIntPtr,
                                             ByVal lParam As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Friend Shared Function GetCurrentThreadId() As UInt32
    End Function

    Friend Const WM_QUIT As Integer = &H12

End Class


<StructLayout(LayoutKind.Sequential)> _
Friend Structure MSG
    Public hWnd As IntPtr
    Public message As UInt32
    Public wParam As IntPtr
    Public lParam As IntPtr
    Public time As UInt32
    Public pt As POINT
End Structure


<StructLayout(LayoutKind.Sequential)> _
Friend Structure POINT
    Public X As Integer
    Public Y As Integer

    Public Sub New(ByVal x As Integer, ByVal y As Integer)
        Me.X = x
        Me.Y = y
    End Sub
End Structure