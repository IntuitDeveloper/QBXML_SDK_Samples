
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Reflection


Friend Class COMRegister

    Public Shared Sub RegisterAsLocalServer(ByVal type As Type)
        COMRegister.CheckForNullType(type, "t")
        Using clsid As RegistryKey = Registry.ClassesRoot.OpenSubKey(
        "CLSID\" & type.GUID.ToString("B"), True)

            clsid.DeleteSubKeyTree("InprocServer32")

            Using subkey As RegistryKey = clsid.CreateSubKey("LocalServer32")
                subkey.SetValue("", Assembly.GetExecutingAssembly.Location,
                                RegistryValueKind.String)
            End Using
        End Using
    End Sub

    Public Shared Sub UnRegisterAsLocalServer(ByVal type As Type)
        COMRegister.CheckForNullType(type, "t")
        Registry.ClassesRoot.DeleteSubKeyTree(("CLSID\" & type.GUID.ToString("B")))
    End Sub


    Private Shared Sub CheckForNullType(ByVal type As Type, ByVal param As String)
        If (type Is Nothing) Then
            Throw New ArgumentException("The CLR typeis not passed", param)
        End If
    End Sub

End Class


Friend Class ComInit

    <DllImport("ole32.dll")>
    Public Shared Function CoInitializeEx(ByVal pvReserved As IntPtr, ByVal dwCoInit As UInt32) As Integer
    End Function

    <DllImport("ole32.dll")>
    Public Shared Sub CoUninitialize()
    End Sub

    <DllImport("ole32.dll")>
    Public Shared Function CoRegisterClassObject(
    ByRef rclsid As Guid, <MarshalAs(UnmanagedType.Interface)> ByVal pUnk As IQBEventHandlerClassFactory,
    ByVal dwClsContext As CLSCTX, ByVal flags As REGCLS, <Out()> ByRef lpdwRegister As UInt32) _
    As Integer
    End Function

    <DllImport("ole32.dll")>
    Public Shared Function CoRevokeClassObject(ByVal dwRegister As UInt32) As UInt32
    End Function

    <DllImport("ole32.dll")>
    Public Shared Function CoResumeClassObjects() As Integer
    End Function

    Public Const IID_IDispatch As String = "00020400-0000-0000-C000-000000000046"

    Public Const IID_IUnknown As String = "00000000-0000-0000-C000-000000000046"

    Public Const CLASS_E_NOAGGREGATION As Integer = &H80040110

    Public Const E_NOINTERFACE As Integer = &H80004002

End Class


<Flags()> _
Friend Enum CLSCTX As UInt32
    INPROC_SERVER = &H1
    INPROC_HANDLER = &H2
    LOCAL_SERVER = &H4
    INPROC_SERVER16 = &H8
    REMOTE_SERVER = &H10
    INPROC_HANDLER16 = &H20
    RESERVED1 = &H40
    RESERVED2 = &H80
    RESERVED3 = &H100
    RESERVED4 = &H200
    NO_CODE_DOWNLOAD = &H400
    RESERVED5 = &H800
    NO_CUSTOM_MARSHAL = &H1000
    ENABLE_CODE_DOWNLOAD = &H2000
    NO_FAILURE_LOG = &H4000
    DISABLE_AAA = &H8000
    ENABLE_AAA = &H10000
    FROM_DEFAULT_CONTEXT = &H20000
    ACTIVATE_32_BIT_SERVER = &H40000
    ACTIVATE_64_BIT_SERVER = &H80000
End Enum

<Flags()> _
Friend Enum REGCLS As UInt32
    SINGLEUSE = 0
    MULTIPLEUSE = 1
    MULTI_SEPARATE = 2
    SUSPENDED = 4
    SURROGATE = 8
End Enum