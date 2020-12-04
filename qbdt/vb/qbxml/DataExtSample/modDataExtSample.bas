Attribute VB_Name = "modDataExtSample"
'----------------------------------------------------------
' Module: modDataExtSample
'
' Description: This module contains the code which creates qbXML
'              messages, exchanges them with QuickBooks, interprets
'              the responses and loads information into form objects.
'
'              The routines in this module are hardcoded to use the
'              OwnerID {E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}.  The
'              string constant cstrOwnerID is defined with this value.
'
' Routines: OpenConnectionBeginSession
'             Opens a connection and begins a sesson with the
'             currently open company file.  If a company isn't open,
'             the routine will display a message and then exit the
'             program.
'
'           GetDataExts
'             Builds a DataExtQueryRq request, sends it to QuickBooks
'             and passes the reply to FillDataExts
'
'           GetCustomFields
'             Builds a DataExtQueryRq request, sends it to QuickBooks
'             and passes the reply to FillDataExts
'
'           FillDataExts
'             Parses a DataExtDefQueryRs response message and fills
'             a passed text box with the data extension definitions
'             returned by the query
'
'           AddDataExtDef
'             Builds a DataExtDefAddRq message, sends it to
'             QuickBooks, parsed the response and if successful calls
'             GetDataExts to add the new data extension definition to
'             the main form, frmDataExtSample
'
'           GetCustomers
'             Builds a CustomerQueryRq message, sends it to
'             QuickBooks and calls FillCustomerListBox
'
'           FillCustomerListBox
'             Parses the passed CustomerQueryRs response and adds
'             sublevel 0 (top level customers, not jobs) to a
'             the listbox passed to the routine.
'
'           CustomersHaveDataExts
'             Calls GetDataExts to fill a text box with the defined
'             data extensions, searches to see if any of them contain
'             the word "Customer" (indicating they are assigned to
'             customer records) and returns True or False depending
'             on the result of the search.
'
'           FillUnusedDataExts
'             Calls GetDataExts to get a list of all data extensions
'             available and makes a list of those assigned for
'             customers records, calls GetUsedDataExts to get the
'             list of data extensions already in use for the passed
'             in customer, compares the lists and puts the unused
'             data extension names and types into the passed in
'             listbox
'
'           GetUsedDataExts
'             Builds a CustomerQueryRq for the passed in customer with
'             OwnerID set to the constant for this module causing the
'             return of data extensions for that customer record to
'             be returned.  It then passes the response to
'             FillUsedDataExts
'
'           FillUsedDataExts
'             Parses a CustomerQueryRq response for a single customer,
'             extracts the returned data extensions (if any) and
'             adds them to a passed in list box with or without their
'             values depending on a passed in boolean
'
'           AddDataExt
'             Using values passed in adds a data extension to a
'             customer record, building the request, parsing the
'             response and displaying a message box with the result
'             and returning a boolean indicating the add success
'
'           ModDataExt
'             Using values passed in modifies a data extension to a
'             customer record, building the request, parsing the
'             response and displaying a message box with the result
'             and returning a boolean indicating the modify success
'
'           EndSessionCloseConnection
'             Calls EndSession and CloseConnection if the boolean
'             booSessionBegun is true.
'
'           PrettyXMLString
'             Adds intentations and newlines to an XML string to make
'             it more attractive for display in a text window.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
'----------------------------------------------------------

  Const cstrOwnerID = "{E09C86CF-9D6E-4EF2-BCBE-4D66B6B0F754}"
  
  Dim booSessionBegun As Boolean
  
  Dim strTicket As String
  
  Public strAddRequest As String
  Public strAddResponse As String
  Public strModRequest As String
  Public strModResponse As String
  Public strQueryRequest As String
  Public strQueryResponse As String
  Public strCustomRequest As String
  Public strCustomResponse As String
  
'Module objects
  Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2
  
  
Public Sub OpenConnectionBeginSession()

  booSessionBegun = False
  
  On Error GoTo Errs
  
  Set qbXMLCOM = New QBXMLRP2Lib.RequestProcessor2

  qbXMLCOM.OpenConnection "", "IDN Data Extension Sample"
  
  strTicket = qbXMLCOM.BeginSession("", QBXMLRP2Lib.qbFileOpenDoNotCare)
  booSessionBegun = True
  
  'Check to make sure that this connection supports at least 2.0
  Dim supportedVersion As String
  supportedVersion = qbXMLLatestVersion(qbXMLCOM, strTicket)
  If (supportedVersion >= "2.0") Then Exit Sub
  
  
  MsgBox "The Data Extension Sample can only run against QuickBooks which support at least version 2.0 of the SDK.  This QuickBooks installation does not."
  End
  
Errs:
  If Err.Number = &H80040416 Then
    MsgBox "You must have QuickBooks running with the company" & vbCrLf & _
           "file open to use this program."
    qbXMLCOM.CloseConnection
    End
  ElseIf Err.Number = &H80040422 Then
    MsgBox "This QuickBooks company file is open in single user mode and" & vbCrLf & _
           "another application is already accessing it.  Please exit the" & vbCrLf & _
           "other application and run this application again."
    qbXMLCOM.CloseConnection
    End
  Else
    MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & _
           ") " & vbCrLf & vbCrLf & Err.Description
    
    If booSessionBegun Then
      qbXMLCOM.EndSession strTicket
    End If
    
    qbXMLCOM.CloseConnection
    End
  End If
End Sub


Public Sub GetDataExts(txtDataExts As TextBox)

Dim strXMLRequest As String
Dim strXMLResponse As String

  strXMLRequest = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""2.0""?>" & _
    "<QBXML>" & _
    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
    "<DataExtDefQueryRq requestID = ""0"">" & _
    "<OwnerID>" & cstrOwnerID & "</OwnerID>" & _
    "</DataExtDefQueryRq>" & _
    "</QBXMLMsgsRq>" & _
    "</QBXML>"
    
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)
  
  strQueryRequest = strXMLRequest
  strQueryResponse = strXMLResponse
  
  FillDataExts strXMLResponse, txtDataExts
End Sub


Public Sub GetCustomFields(txtCustomFields As TextBox)

Dim strXMLRequest As String
Dim strXMLResponse As String

  strXMLRequest = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""2.0""?>" & _
    "<QBXML>" & _
    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
    "<DataExtDefQueryRq requestID = ""0"">" & _
    "<OwnerID>0</OwnerID>" & _
    "</DataExtDefQueryRq>" & _
    "</QBXMLMsgsRq>" & _
    "</QBXML>"
    
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)

  strCustomRequest = strXMLRequest
  strCustomResponse = strXMLResponse
  
  FillDataExts strXMLResponse, txtCustomFields
End Sub


Public Sub FillDataExts(strXMLResponse As String, _
                        txtDataExts As TextBox)

'Set up a DOM document object to load the response into
  Dim xmlDataExtDefResponse As New MSXML2.DOMDocument
  Set xmlDataExtDefResponse = CreateObject("MSXML2.DOMDocument")

'Clear the text box
  txtDataExts.Text = ""
  txtDataExts.Refresh
  
'Parse the response XML
  xmlDataExtDefResponse.async = False
  xmlDataExtDefResponse.loadXML (strXMLResponse)

'Get the status for our query request
  Dim nodeDataExtDefQueryRs As IXMLDOMNodeList
  Set nodeDataExtDefQueryRs = xmlDataExtDefResponse.getElementsByTagName("DataExtDefQueryRs")
  
  Dim rsStatusAttr        As IXMLDOMNamedNodeMap
  Set rsStatusAttr = nodeDataExtDefQueryRs.Item(0).Attributes
  Dim strQueryStatus As String
  strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

'If the status is bad, report it to the user
  If strQueryStatus <> "0" Then
    If strQueryStatus = "1" Then
      txtDataExts.Text = "NO DATA EXTENSIONS DEFINED FOR THIS COMPANY FILE"
    Else
      MsgBox "FillDataExts unexpexcted Error - " & rsStatusAttr.getNamedItem("statusMessage").nodeValue
    End If
    Exit Sub
  End If
  
'Parse the query response and add the DataExtDefs to the DataExtDef text box
  Dim DataExtDefQueryRsNode As IXMLDOMNode
  Set DataExtDefQueryRsNode = nodeDataExtDefQueryRs.Item(0)
  
  Dim DataExtDefNodeList    As IXMLDOMNodeList
  Set DataExtDefNodeList = xmlDataExtDefResponse.getElementsByTagName("DataExtDefRet")
  
  Dim numItems            As Long
  numItems = DataExtDefNodeList.length
  
'Declare the XML objects outside of the loop
  Dim ReturnElement      As IXMLDOMElement
  Dim itemNode           As IXMLDOMNode
  Dim objectNode         As IXMLDOMNode
  Dim AssignToObjectList As IXMLDOMNodeList
  Dim i                  As Integer
  Dim j                  As Integer
  Dim numObjects         As Long
  Dim strDisplayLine     As String
  
  For i = 0 To numItems - 1
    Set itemNode = DataExtDefNodeList.Item(i)
    strDisplayLine = ""
    Set ReturnElement = itemNode.selectSingleNode("DataExtName")
    strDisplayLine = ReturnElement.Text
    Set ReturnElement = itemNode.selectSingleNode("DataExtType")
    strDisplayLine = strDisplayLine & "  |  " & ReturnElement.Text & "  |  "
    
    'Parse the Data Extension return object Assigned Objects to the DataExtDef text box
    Set AssignToObjectList = itemNode.selectNodes("AssignToObject")
    numObjects = AssignToObjectList.length
    For j = 0 To numObjects - 1
      Set objectNode = AssignToObjectList.Item(j)
      If j > 0 Then
        strDisplayLine = strDisplayLine & ", " & objectNode.Text
      Else
        strDisplayLine = strDisplayLine & objectNode.Text
      End If
    Next
    
    txtDataExts.Text = txtDataExts.Text & strDisplayLine & vbCrLf
  Next
End Sub


Public Sub AddDataExtDef _
    (strDataExtName As String, _
     strDataExtType As String, _
     strObjects As String)

  Dim strXMLRequest As String
  Dim strXMLResponse As String
  
'Create a DOM document object for creating our request.
  Dim xmlDataExtDefAdd As New MSXML2.DOMDocument
  Set xmlDataExtDefAdd = CreateObject("MSXML2.DOMDocument")

'Add the QBXML aggregate
  Dim rootElement As IXMLDOMNode
  Set rootElement = xmlDataExtDefAdd.createElement("QBXML")
  xmlDataExtDefAdd.appendChild rootElement
  
'Add the QBXMLMsgsRq aggregate
  Dim QBXMLMsgsRqNode As IXMLDOMNode
  Set QBXMLMsgsRqNode = xmlDataExtDefAdd.createElement("QBXMLMsgsRq")
  rootElement.appendChild QBXMLMsgsRqNode
  
'Set the QBXMLMsgsRq onError attribute to continueOnError
  Dim onErrorAttr As IXMLDOMAttribute
  Set onErrorAttr = xmlDataExtDefAdd.createAttribute("onError")
  onErrorAttr.Text = "stopOnError"
  QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
'Add the DataExtDefAddRq aggregate
  Dim DataExtDefAddRqNode As IXMLDOMNode
  Set DataExtDefAddRqNode = _
    xmlDataExtDefAdd.createElement("DataExtDefAddRq")
  QBXMLMsgsRqNode.appendChild DataExtDefAddRqNode
  
'Add the DataExtDefAdd aggregate
  Dim DataExtDefAddNode As IXMLDOMNode
  Set DataExtDefAddNode = _
    xmlDataExtDefAdd.createElement("DataExtDefAdd")
  DataExtDefAddRqNode.appendChild DataExtDefAddNode
  
'Add the OwnerID element
  Dim OwnerIDElement As IXMLDOMElement
  Set OwnerIDElement = xmlDataExtDefAdd.createElement("OwnerID")
  OwnerIDElement.Text = cstrOwnerID
  DataExtDefAddNode.appendChild OwnerIDElement
  
'Add the DataExtName element
  Dim DataExtNameElement As IXMLDOMElement
  Set DataExtNameElement = xmlDataExtDefAdd.createElement("DataExtName")
  DataExtNameElement.Text = strDataExtName
  DataExtDefAddNode.appendChild DataExtNameElement
  
'Add the DataExtType element
  Dim DataExtTypeElement As IXMLDOMElement
  Set DataExtTypeElement = xmlDataExtDefAdd.createElement("DataExtType")
  DataExtTypeElement.Text = strDataExtType
  DataExtDefAddNode.appendChild DataExtTypeElement
  
'Add the AssignToObject elements
  Dim strAssignToObject() As String
  Dim i As Integer
  Dim AssignToObjectElement As IXMLDOMElement
  
  strAssignToObject = Split(strObjects)
  For i = LBound(strAssignToObject) To UBound(strAssignToObject)
    Set AssignToObjectElement = xmlDataExtDefAdd.createElement("AssignToObject")
    AssignToObjectElement.Text = strAssignToObject(i)
    DataExtDefAddNode.appendChild AssignToObjectElement
  Next
  
'Set the requestID attribute to 0
  Dim requestIDAttr As IXMLDOMAttribute
  Set requestIDAttr = xmlDataExtDefAdd.createAttribute("requestID")
  requestIDAttr.Text = "0"
  DataExtDefAddRqNode.Attributes.setNamedItem requestIDAttr
  
'We're adding the prolog using text strings
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""2.0""?>" & _
    rootElement.xml

On Error GoTo Errs

'Send the request to QuickBooks
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)

'Parse the response and check for errors
'Set up a DOM document object to load the response into
  Dim xmlDataExtDefAddResponse As New MSXML2.DOMDocument
  Set xmlDataExtDefAddResponse = CreateObject("MSXML2.DOMDocument")

'Parse the response XML
  xmlDataExtDefAddResponse.async = False
  xmlDataExtDefAddResponse.loadXML (strXMLResponse)

'Get the status for our query request
  Dim nodeDataExtDefAddRs As IXMLDOMNodeList
  Set nodeDataExtDefAddRs = xmlDataExtDefAddResponse.getElementsByTagName("DataExtDefAddRs")
  
  Dim rsStatusAttr        As IXMLDOMNamedNodeMap
  Set rsStatusAttr = nodeDataExtDefAddRs.Item(0).Attributes
  Dim strAddStatus As String
  strAddStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

'If the status is bad, report it to the user
  If strAddStatus <> "0" Then
    MsgBox rsStatusAttr.getNamedItem("statusMessage").nodeValue
    Exit Sub
  End If

'Since we added a new data extension, update the data extensions text box in the main form
  GetDataExts frmDataExtSample.txtDataExts

  strAddRequest = strXMLRequest
  strAddResponse = strXMLResponse

  Exit Sub

Errs:
  MsgBox Str(Err.Number) & "  -  " & Err.Description
End Sub


Public Sub GetCustomers(lstCustomers As ListBox)

Dim strXMLRequest As String
Dim strXMLResponse As String

  strXMLRequest = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""4.0""?>" & _
    "<QBXML>" & _
    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
    "<CustomerQueryRq requestID = ""0"">" & _
    "<IncludeRetElement>FullName</IncludeRetElement>" & _
    "</CustomerQueryRq>" & _
    "</QBXMLMsgsRq>" & _
    "</QBXML>"
    
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)
  
  FillCustomerListBox strXMLResponse, lstCustomers
End Sub


Public Sub FillCustomerListBox(strXMLResponse As String, _
                               lstCustomers As ListBox)

'Set up a DOM document object to load the response into
  Dim xmlCustomerResponse As New MSXML2.DOMDocument
  Set xmlCustomerResponse = CreateObject("MSXML2.DOMDocument")

'Clear the list box
  lstCustomers.Clear
  
'Parse the response XML
  xmlCustomerResponse.async = False
  xmlCustomerResponse.loadXML (strXMLResponse)

'Get the status for our query request
  Dim nodeCustomerQueryRs As IXMLDOMNodeList
  Set nodeCustomerQueryRs = xmlCustomerResponse.getElementsByTagName("CustomerQueryRs")
  
  Dim rsStatusAttr        As IXMLDOMNamedNodeMap
  Set rsStatusAttr = nodeCustomerQueryRs.Item(0).Attributes
  Dim strQueryStatus As String
  strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

'If the status is bad, report it to the user
  If strQueryStatus <> "0" Then
    MsgBox rsStatusAttr.getNamedItem("statusMessage").nodeValue
    Exit Sub
  End If
  
'Parse the query response and add the Customers to the Customer list box
  Dim CustomerQueryRsNode As IXMLDOMNode
  Set CustomerQueryRsNode = nodeCustomerQueryRs.Item(0)
  
  Dim CustomerNodeList    As IXMLDOMNodeList
  Set CustomerNodeList = xmlCustomerResponse.getElementsByTagName("CustomerRet")
  
  Dim numItems            As Long
  numItems = CustomerNodeList.length
  
'Declare the XML objects outside of the loop
  Dim ReturnElement     As IXMLDOMElement
  Dim itemNode          As IXMLDOMNode
  Dim i                 As Integer
  Dim strDisplayLine    As String
  For i = 0 To numItems - 1
    Set itemNode = CustomerNodeList.Item(i)
    'The following two lines are no longer necessary because we
    'limited the CustomerQuery to return only the FullName.
    'Set ReturnElement = itemNode.selectSingleNode("Sublevel")
    'If ReturnElement.Text = "0" Then
      Set ReturnElement = itemNode.selectSingleNode("FullName")
      lstCustomers.AddItem ReturnElement.Text
   ' End If
  Next
End Sub


Public Function CustomersHaveDataExts() As Boolean

  GetDataExts frmDataExtSample.txtUtility
  If InStr(frmDataExtSample.txtUtility.Text, "Customer") Then
    CustomersHaveDataExts = True
  Else
    CustomersHaveDataExts = False
  End If
End Function


Public Sub FillUnusedDataExts(strCustomer As String, _
                              lstUnusedDataExtDefs As ListBox)

  Dim strDataExts() As String
  Dim i             As Integer
  Dim dummy As String
  
  lstUnusedDataExtDefs.Clear
  frmAddDataExtension.lstAvailableDataExts.Clear
  frmAddDataExtension.lstUsedDataExts.Clear
  
  GetDataExts frmDataExtSample.txtUtility
  strDataExts = Split(frmDataExtSample.txtUtility.Text, vbCrLf)
  For i = LBound(strDataExts) To UBound(strDataExts) - 1
    If InStr(strDataExts(i), "Customer") > 0 Then
      frmAddDataExtension.lstAvailableDataExts.AddItem _
        Left(strDataExts(i), InStr(strDataExts(i), "|  Customer") - 3)
    End If
  Next
  
  GetUsedCustomerDataExts strCustomer, frmAddDataExtension.lstUsedDataExts, False
  
  'If the list counts are equal then the given customer has values
  'assigned for all of their possible data extensions.
  If frmAddDataExtension.lstUsedDataExts.ListCount = _
     frmAddDataExtension.lstAvailableDataExts.ListCount Then
    lstUnusedDataExtDefs.AddItem _
    "All Data Extensions for " & strCustomer & " are already assigned values"
    Exit Sub
  End If
  
  'If none are used, they are all available
  If frmAddDataExtension.lstUsedDataExts.ListCount = 0 Then
    For i = 0 To frmAddDataExtension.lstAvailableDataExts.ListCount - 1
      lstUnusedDataExtDefs.AddItem _
        frmAddDataExtension.lstAvailableDataExts.List(i)
    Next
    Exit Sub
  End If
  
  'Some are used, so lets figure out which ones
  Dim j As Integer
  
  j = 0
  For i = 0 To frmAddDataExtension.lstUsedDataExts.ListCount - 1
    Do While Left(frmAddDataExtension.lstAvailableDataExts.List(j), _
                  InStr(frmAddDataExtension.lstAvailableDataExts.List(j), "  |") - 1) < _
             frmAddDataExtension.lstUsedDataExts.List(i)
      lstUnusedDataExtDefs.AddItem frmAddDataExtension.lstAvailableDataExts.List(j)
      j = j + 1
    Loop
    j = j + 1
  Next

'List the remaining avaliable data extension definitions if any
  Do While j <= frmAddDataExtension.lstAvailableDataExts.ListCount - 1
    lstUnusedDataExtDefs.AddItem frmAddDataExtension.lstAvailableDataExts.List(j)
    j = j + 1
  Loop
End Sub


Public Sub GetUsedCustomerDataExts(strCustomer As String, _
                                   lstUsedDataExts As ListBox, _
                                   booIncludeValues As Boolean)

Dim strXMLRequest As String
Dim strXMLResponse As String

  strXMLRequest = "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""2.0""?>" & _
    "<QBXML>" & _
    "<QBXMLMsgsRq onError = ""continueOnError"">" & _
    "<CustomerQueryRq requestID = ""0"">" & _
    "<FullName>" & strCustomer & "</FullName>" & _
    "<OwnerID>" & cstrOwnerID & "</OwnerID>" & _
    "</CustomerQueryRq>" & _
    "</QBXMLMsgsRq>" & _
    "</QBXML>"
    
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)

  strQueryRequest = strXMLRequest
  strQueryResponse = strXMLResponse

  FillUsedDataExts strXMLResponse, lstUsedDataExts, booIncludeValues
End Sub


Public Sub FillUsedDataExts(strXMLResponse As String, _
                            lstUsedDataExts As ListBox, _
                            booIncludeValues As Boolean)

'Set up a DOM document object to load the response into
  Dim xmlCustomerResponse As New MSXML2.DOMDocument
  Set xmlCustomerResponse = CreateObject("MSXML2.DOMDocument")

'Clear the list box
  lstUsedDataExts.Clear
  
'Parse the response XML
  xmlCustomerResponse.async = False
  xmlCustomerResponse.loadXML (strXMLResponse)

'Get the status for our query request
  Dim nodeCustomerQueryRs As IXMLDOMNodeList
  Set nodeCustomerQueryRs = xmlCustomerResponse.getElementsByTagName("CustomerQueryRs")
  
  Dim rsStatusAttr        As IXMLDOMNamedNodeMap
  Set rsStatusAttr = nodeCustomerQueryRs.Item(0).Attributes
  Dim strQueryStatus As String
  strQueryStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

'If the status is bad, report it to the user
  If strQueryStatus <> "0" Then
    MsgBox rsStatusAttr.getNamedItem("statusMessage").nodeValue
    Exit Sub
  End If
  
'Parse the query response and add the Customers to the Customer list box
  Dim CustomerQueryRsNode As IXMLDOMNode
  Set CustomerQueryRsNode = nodeCustomerQueryRs.Item(0)
  
'We're only expecting one CustomerRet record, so get the node.
  Dim CustomerNode    As IXMLDOMNode
  Set CustomerNode = nodeCustomerQueryRs.Item(0).selectSingleNode("CustomerRet")

  Dim DataExtsNodeList As IXMLDOMNodeList
  Set DataExtsNodeList = CustomerNode.selectNodes("DataExtRet")
  
  If DataExtsNodeList.length = 0 Then Exit Sub
  
  Dim numItems            As Long
  numItems = DataExtsNodeList.length
  
'Declare the XML objects outside of the loop
  Dim ReturnElement     As IXMLDOMElement
  Dim itemNode          As IXMLDOMNode
  Dim i                 As Integer
  Dim strDisplayLine    As String
  For i = 0 To numItems - 1
    strDisplayLine = ""
    Set itemNode = DataExtsNodeList.Item(i)
    Set ReturnElement = itemNode.selectSingleNode("DataExtName")
    strDisplayLine = ReturnElement.Text
    If booIncludeValues Then
      Set ReturnElement = itemNode.selectSingleNode("DataExtValue")
      strDisplayLine = strDisplayLine & "  =  " & ReturnElement.Text
    End If
    lstUsedDataExts.AddItem strDisplayLine
  Next
End Sub


Public Function AddDataExt(strCustomer As String, _
                           strDataExtName As String, _
                           strDataExtValue As String) As Boolean

  Dim strXMLRequest As String
  Dim strXMLResponse As String
  
'Create a DOM document object for creating our request.
  Dim xmlDataExtAdd As New MSXML2.DOMDocument
  Set xmlDataExtAdd = CreateObject("MSXML2.DOMDocument")

'Add the QBXML aggregate
  Dim rootElement As IXMLDOMNode
  Set rootElement = xmlDataExtAdd.createElement("QBXML")
  xmlDataExtAdd.appendChild rootElement
  
'Add the QBXMLMsgsRq aggregate
  Dim QBXMLMsgsRqNode As IXMLDOMNode
  Set QBXMLMsgsRqNode = xmlDataExtAdd.createElement("QBXMLMsgsRq")
  rootElement.appendChild QBXMLMsgsRqNode
  
'Set the QBXMLMsgsRq onError attribute to continueOnError
  Dim onErrorAttr As IXMLDOMAttribute
  Set onErrorAttr = xmlDataExtAdd.createAttribute("onError")
  onErrorAttr.Text = "stopOnError"
  QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
'Add the DataExtAddRq aggregate
  Dim DataExtAddRqNode As IXMLDOMNode
  Set DataExtAddRqNode = _
    xmlDataExtAdd.createElement("DataExtAddRq")
  QBXMLMsgsRqNode.appendChild DataExtAddRqNode
  
'Add the DataExtAdd aggregate
  Dim DataExtAddNode As IXMLDOMNode
  Set DataExtAddNode = _
    xmlDataExtAdd.createElement("DataExtAdd")
  DataExtAddRqNode.appendChild DataExtAddNode
  
'Add the OwnerID element
  Dim OwnerIDElement As IXMLDOMElement
  Set OwnerIDElement = xmlDataExtAdd.createElement("OwnerID")
  OwnerIDElement.Text = cstrOwnerID
  DataExtAddNode.appendChild OwnerIDElement
  
'Add the DataExtName element
  Dim DataExtNameElement As IXMLDOMElement
  Set DataExtNameElement = xmlDataExtAdd.createElement("DataExtName")
  DataExtNameElement.Text = strDataExtName
  DataExtAddNode.appendChild DataExtNameElement
  
'Add the ListDataExtType element
  Dim ListDataExtTypeElement As IXMLDOMElement
  Set ListDataExtTypeElement = xmlDataExtAdd.createElement("ListDataExtType")
  ListDataExtTypeElement.Text = "Customer"
  DataExtAddNode.appendChild ListDataExtTypeElement
  
'Add the ListObjRef aggregate
  Dim ListObjRefNode As IXMLDOMNode
  Set ListObjRefNode = xmlDataExtAdd.createElement("ListObjRef")
  DataExtAddNode.appendChild ListObjRefNode
  
'Add the FullName element
  Dim FullNameElement As IXMLDOMElement
  Set FullNameElement = xmlDataExtAdd.createElement("FullName")
  FullNameElement.Text = strCustomer
  ListObjRefNode.appendChild FullNameElement
  
'Add the DataExtValue element
  Dim DataExtValueElement As IXMLDOMElement
  Set DataExtValueElement = xmlDataExtAdd.createElement("DataExtValue")
  DataExtValueElement.Text = strDataExtValue
  DataExtAddNode.appendChild DataExtValueElement
  
'Set the requestID attribute to 0
  Dim requestIDAttr As IXMLDOMAttribute
  Set requestIDAttr = xmlDataExtAdd.createAttribute("requestID")
  requestIDAttr.Text = "0"
  DataExtAddRqNode.Attributes.setNamedItem requestIDAttr
  
'We're adding the prolog using text strings
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""2.0""?>" & _
    rootElement.xml

On Error GoTo Errs

'Send the request to QuickBooks
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)

'Parse the response and check for errors
'Set up a DOM document object to load the response into
  Dim xmlDataExtAddResponse As New MSXML2.DOMDocument
  Set xmlDataExtAddResponse = CreateObject("MSXML2.DOMDocument")

'Parse the response XML
  xmlDataExtAddResponse.async = False
  xmlDataExtAddResponse.loadXML (strXMLResponse)

'Get the status for our query request
  Dim nodeDataExtAddRs As IXMLDOMNodeList
  Set nodeDataExtAddRs = xmlDataExtAddResponse.getElementsByTagName("DataExtAddRs")
  
  Dim rsStatusAttr        As IXMLDOMNamedNodeMap
  Set rsStatusAttr = nodeDataExtAddRs.Item(0).Attributes
  Dim strAddStatus As String
  strAddStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue

  strAddRequest = strXMLRequest
  strAddResponse = strXMLResponse

'If the status is bad, report it to the user
  If strAddStatus <> "0" Then
    MsgBox "AddDataExt - " & rsStatusAttr.getNamedItem("statusMessage").nodeValue
    AddDataExt = False
    Exit Function
  End If

  MsgBox "Successfully added Data Extension " & strDataExtName & _
         " with value " & strDataExtValue & " to customer " & _
         strCustomer
  AddDataExt = True
  Exit Function

Errs:
  MsgBox "AddDataExt - " & Str(Err.Number) & "  -  " & Err.Description
  AddDataExt = False
End Function


Public Function ModDataExt(strCustomer As String, _
                           strDataExtName As String, _
                           strDataExtValue As String) As Boolean

  Dim strXMLRequest As String
  Dim strXMLResponse As String
  
'Create a DOM document object for creating our request.
  Dim xmlDataExtMod As New MSXML2.DOMDocument
  Set xmlDataExtMod = CreateObject("MSXML2.DOMDocument")

'Add the QBXML aggregate
  Dim rootElement As IXMLDOMNode
  Set rootElement = xmlDataExtMod.createElement("QBXML")
  xmlDataExtMod.appendChild rootElement
  
'Add the QBXMLMsgsRq aggregate
  Dim QBXMLMsgsRqNode As IXMLDOMNode
  Set QBXMLMsgsRqNode = xmlDataExtMod.createElement("QBXMLMsgsRq")
  rootElement.appendChild QBXMLMsgsRqNode
  
'Set the QBXMLMsgsRq onError attribute to continueOnError
  Dim onErrorAttr As IXMLDOMAttribute
  Set onErrorAttr = xmlDataExtMod.createAttribute("onError")
  onErrorAttr.Text = "stopOnError"
  QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
'Add the DataExtModRq aggregate
  Dim DataExtModRqNode As IXMLDOMNode
  Set DataExtModRqNode = _
    xmlDataExtMod.createElement("DataExtModRq")
  QBXMLMsgsRqNode.appendChild DataExtModRqNode
  
'Add the DataExtMod aggregate
  Dim DataExtModNode As IXMLDOMNode
  Set DataExtModNode = _
    xmlDataExtMod.createElement("DataExtMod")
  DataExtModRqNode.appendChild DataExtModNode
  
'Add the OwnerID element
  Dim OwnerIDElement As IXMLDOMElement
  Set OwnerIDElement = xmlDataExtMod.createElement("OwnerID")
  OwnerIDElement.Text = cstrOwnerID
  DataExtModNode.appendChild OwnerIDElement
  
'Add the DataExtName element
  Dim DataExtNameElement As IXMLDOMElement
  Set DataExtNameElement = xmlDataExtMod.createElement("DataExtName")
  DataExtNameElement.Text = strDataExtName
  DataExtModNode.appendChild DataExtNameElement
  
'Add the ListDataExtType element
  Dim ListDataExtTypeElement As IXMLDOMElement
  Set ListDataExtTypeElement = xmlDataExtMod.createElement("ListDataExtType")
  ListDataExtTypeElement.Text = "Customer"
  DataExtModNode.appendChild ListDataExtTypeElement
  
'Add the ListObjRef aggregate
  Dim ListObjRefNode As IXMLDOMNode
  Set ListObjRefNode = xmlDataExtMod.createElement("ListObjRef")
  DataExtModNode.appendChild ListObjRefNode
  
'Add the FullName element
  Dim FullNameElement As IXMLDOMElement
  Set FullNameElement = xmlDataExtMod.createElement("FullName")
  FullNameElement.Text = strCustomer
  ListObjRefNode.appendChild FullNameElement
  
'Add the DataExtValue element
  Dim DataExtValueElement As IXMLDOMElement
  Set DataExtValueElement = xmlDataExtMod.createElement("DataExtValue")
  DataExtValueElement.Text = strDataExtValue
  DataExtModNode.appendChild DataExtValueElement
  
'Set the requestID attribute to 0
  Dim requestIDAttr As IXMLDOMAttribute
  Set requestIDAttr = xmlDataExtMod.createAttribute("requestID")
  requestIDAttr.Text = "0"
  DataExtModRqNode.Attributes.setNamedItem requestIDAttr
  
'We're adding the prolog using text strings
  strXMLRequest = _
    "<?xml version=""1.0"" ?>" & _
    "<?qbxml version=""2.0""?>" & _
    rootElement.xml

On Error GoTo Errs

'Send the request to QuickBooks
  strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)

'Parse the response and check for errors
'Set up a DOM document object to load the response into
  Dim xmlDataExtModResponse As New MSXML2.DOMDocument
  Set xmlDataExtModResponse = CreateObject("MSXML2.DOMDocument")

'Parse the response XML
  xmlDataExtModResponse.async = False
  xmlDataExtModResponse.loadXML (strXMLResponse)

'Get the status for our query request
  Dim nodeDataExtModRs As IXMLDOMNodeList
  Set nodeDataExtModRs = xmlDataExtModResponse.getElementsByTagName("DataExtModRs")
  
  Dim rsStatusAttr        As IXMLDOMNamedNodeMap
  Set rsStatusAttr = nodeDataExtModRs.Item(0).Attributes
  Dim strModStatus As String
  strModStatus = rsStatusAttr.getNamedItem("statusCode").nodeValue
  
  strModRequest = strXMLRequest
  strModResponse = strXMLResponse

'If the status is bad, report it to the user
  If strModStatus <> "0" Then
    MsgBox "ModDataExt - " & rsStatusAttr.getNamedItem("statusMessage").nodeValue
    ModDataExt = False
    Exit Function
  End If

  MsgBox "Successfully added Data Extension " & strDataExtName & _
         " with value " & strDataExtValue & " to customer " & _
         strCustomer
  ModDataExt = True
  Exit Function

Errs:
  MsgBox "ModDataExt - " & Str(Err.Number) & "  -  " & Err.Description
  ModDataExt = False
End Function


Public Sub EndSessionCloseConnection()
  If booSessionBegun Then
    qbXMLCOM.EndSession strTicket
    qbXMLCOM.CloseConnection
  End If
End Sub

Public Function PrettyXMLString(InXMLString As String) As String
                                       
  Dim SplitXMLString() As String
  Dim IndentString As String
  Dim OutString As String
  Dim XMLString As String
  Dim XMLStringLength As Long
  Dim SplitIndex
  
' Preserve the passed in string
  XMLString = InXMLString
  
  IndentString = Empty
  OutString = Empty
  
'Remove the linefeeds from the XML string
  XMLString = Replace(XMLString, vbLf, vbNullString)
  
  SplitXMLString = Split(XMLString, "<")
  
'We're expecting the first character of the XML string to be "<"
'which result in an empty first array element, so skip it.
  SplitIndex = LBound(SplitXMLString) + 1
  
  Do
    If Left(SplitXMLString(SplitIndex), 1) = "/" Then
      IndentString = Left(IndentString, Len(IndentString) - 3)
      OutString = OutString & IndentString & "<" & _
                      SplitXMLString(SplitIndex) & vbCrLf
      SplitIndex = SplitIndex + 1
    ElseIf Left(SplitXMLString(SplitIndex + 1), 1) = "/" Then
      If InStr(1, Left(SplitXMLString(SplitIndex), _
               InStr(1, SplitXMLString(SplitIndex), ">")), _
                " ") > 0 Then
        OutString = OutString & IndentString & "<" & _
                        SplitXMLString(SplitIndex) & vbCrLf
        SplitIndex = SplitIndex + 1
      Else
        OutString = OutString & IndentString & "<" & _
                        SplitXMLString(SplitIndex) & "<" & _
                        SplitXMLString(SplitIndex + 1) & vbCrLf
        SplitIndex = SplitIndex + 2
      End If
    Else
      OutString = OutString & IndentString & "<" & _
                      SplitXMLString(SplitIndex) & vbCrLf
      If SplitXMLString(SplitIndex) <> "?xml version=""1.0"" ?>" And _
         SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' >" And _
         SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.1//EN' >" And _
         SplitXMLString(SplitIndex) <> "!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 2.0//EN' >" And _
         SplitXMLString(SplitIndex) <> "?qbxml version=""2.0""?>" And _
         Right(SplitXMLString(SplitIndex), 2) <> "/>" Then
        IndentString = IndentString & "   "
      End If
      SplitIndex = SplitIndex + 1
    End If
  Loop Until SplitIndex >= UBound(SplitXMLString)
  
  If Left(SplitXMLString(UBound(SplitXMLString)), 1) = "/" Then
    IndentString = Left(IndentString, Len(IndentString) - 3)
  End If
  
  OutString = OutString & IndentString & "<" & _
                  SplitXMLString(UBound(SplitXMLString))

  PrettyXMLString = OutString
  
End Function

Function qbXMLLatestVersion(rp As RequestProcessor2, ticket As String) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = rp.QBXMLVersionsForSession(ticket)
    
    'Instead we use HostQuery
    'Create a DOM document object for creating our request.
    Dim xml As New DOMDocument

    'Create the QBXML aggregate
    Dim rootElement As IXMLDOMNode
    Set rootElement = xml.createElement("QBXML")
    xml.appendChild rootElement
  
    'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
    Dim QBXMLMsgsRqNode As IXMLDOMNode
    Set QBXMLMsgsRqNode = xml.createElement("QBXMLMsgsRq")
    rootElement.appendChild QBXMLMsgsRqNode

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If we were writing a real application this is where we would add
    'a newMessageSetID so we could perform error recovery.  Any time a
    'request contains an add, delete, modify or void request developers
    'should use the error recovery mechanisms.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Set the QBXMLMsgsRq onError attribute to continueOnError
    Dim onErrorAttr As IXMLDOMAttribute
    Set onErrorAttr = xml.createAttribute("onError")
    onErrorAttr.Text = "stopOnError"
    QBXMLMsgsRqNode.Attributes.setNamedItem onErrorAttr
  
    'Add the InvoiceAddRq aggregate to QBXMLMsgsRq aggregate
    Dim HostQuery As IXMLDOMNode
    Set HostQuery = xml.createElement("HostQueryRq")
    QBXMLMsgsRqNode.appendChild HostQuery
    
    Dim strXMLRequest As String
    strXMLRequest = _
        "<?xml version=""1.0"" ?>" & _
        "<!DOCTYPE QBXML PUBLIC '-//INTUIT//DTD QBXML QBD 1.0//EN' 'http://developer.intuit.com'>" _
        & rootElement.xml

    Dim strXMLResponse As String
    strXMLResponse = rp.ProcessRequest(ticket, strXMLRequest)
    Dim QueryResponse As New DOMDocument

    'Parse the response XML
    QueryResponse.async = False
    QueryResponse.loadXML (strXMLResponse)

    Dim supportedVersions As IXMLDOMNodeList
    Set supportedVersions = QueryResponse.getElementsByTagName("SupportedQBXMLVersion")
    
    Dim VersNode As IXMLDOMNode
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.length - 1
        Set VersNode = supportedVersions.Item(i)
        vers = VersNode.firstChild.Text
        If (vers > LastVers) Then
            LastVers = vers
            qbXMLLatestVersion = VersNode.firstChild.Text
        End If
    Next i
End Function



