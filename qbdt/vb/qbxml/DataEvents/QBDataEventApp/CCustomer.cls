VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' this is just a trivial little object (note all variables are public so we don't even
' have to bother with Property Let/Get methods) to hold the information we care about
' in a customer -- kind of simulates a database record.
'

Public name As String
Public Phone As String
Public Email As String
Public Balance As String
Public index As String
Public ListID As String

Sub Init()
    name = ""
    Phone = ""
    Email = ""
    Balance = "0.00"
    index = ""
    ListID = ""
End Sub

