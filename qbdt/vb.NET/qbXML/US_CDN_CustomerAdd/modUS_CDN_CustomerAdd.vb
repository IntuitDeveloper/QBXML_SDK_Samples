Option Strict Off
Option Explicit On
Module modUS_CDN_CustomerAdd
    '-----------------------------------------------------------
    ' Module: modUS_CDN_CustomerAdd
    '
    ' Description: This module contains the routines which interact
    '              with QuickBooks.
    '
    ' Routines: StartQuickBooksSession
    '             Calls OpenConnection and BeginSession to enable this
    '             sample to communicate with QuickBooks
    '
    '           GetQBNationality
    '             Returns the nationality of the QuickBooks connected to
    '
    '           AddCustomer
    '             Builds a message to add a customer to QuickBooks,
    '             sends it to QB and parses the response to determine if
    '             the customer was added successfully
    '
    '           EndSessionCloseConnection
    '             Ends the QuickBooks session and closes the connection
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------

    'Constants
    Const constCanadianVersion As String = "CA2.0"
	
	'Global Module Variables
	Dim booSessionBegun As Boolean
	
	Dim strTicket As String

    'Module objects
    Dim qbXMLCOM As QBXMLRP2Lib.RequestProcessor2

    Public Sub StartQuickBooksSession()
		qbXMLCOM = New QBXMLRP2Lib.RequestProcessor2
		
		booSessionBegun = False
		
		On Error GoTo Errs
		
		qbXMLCOM.OpenConnection("", "IDN US / Canadian Customer Add")
		
		strTicket = qbXMLCOM.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)
		booSessionBegun = True
		
		Exit Sub
		
Errs: 
		If Err.Number = &H80040416 Then
			MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
			qbXMLCOM.CloseConnection()
			End
		ElseIf Err.Number = &H80040422 Then 
			MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
			qbXMLCOM.CloseConnection()
			End
		Else
			MsgBox("Unable to connect to QuickBooks, HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description)
			
			If booSessionBegun Then
				qbXMLCOM.EndSession(strTicket)
			End If
			
			qbXMLCOM.CloseConnection()
			End
		End If
	End Sub
	
	Public Function GetQBNationality() As String
		
		Dim strVersions() As String
		
		'Get the support versions to find out if this is US or Canadian QB
		strVersions = qbXMLCOM.QBXMLVersionsForSession(strTicket)
		
		'Default to US and set Canadian if it is
		GetQBNationality = "US"
		
		Dim i As Short
		For i = 0 To UBound(strVersions)
			If strVersions(i) = constCanadianVersion Then
				GetQBNationality = "Canadian"
			End If
		Next 
	End Function
	
	Public Function AddCustomer(ByRef strNationality As String, ByRef strCustomerName As String, ByRef strAddr1 As String, ByRef strAddr2 As String, ByRef strAddr3 As String, ByRef strAddr4 As String, ByRef strCity As String, ByRef strStateProvince As String, ByRef strPostalCode As String) As Boolean
		
		Dim strXMLRequest As String
		Dim strXMLResponse As String

        'Create a DOM document object for creating our request.
        Dim xmlCustomerAdd As Object
        xmlCustomerAdd = CreateObject("MSXML2.DOMDocument")

        'Create the QBXML aggregate
        Dim rootElement As MSXML2.IXMLDOMNode
		rootElement = xmlCustomerAdd.createElement("QBXML")
		xmlCustomerAdd.appendChild(rootElement)
		
		'Add the QBXMLMsgsRq aggregate to the QBXML aggregate
		Dim QBXMLMsgsRqNode As MSXML2.IXMLDOMNode
		QBXMLMsgsRqNode = xmlCustomerAdd.createElement("QBXMLMsgsRq")
		rootElement.appendChild(QBXMLMsgsRqNode)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'If we were writing a real application this is where we would add
		'a newMessageSetID so we could perform error recovery.  Any time a
		'request contains an add, delete, modify or void request developers
		'should use the error recovery mechanisms.
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Set the QBXMLMsgsRq onError attribute to continueOnError
		Dim onErrorAttr As MSXML2.IXMLDOMAttribute
		onErrorAttr = xmlCustomerAdd.createAttribute("onError")
		onErrorAttr.Text = "stopOnError"
        QBXMLMsgsRqNode.attributes.setNamedItem(onErrorAttr)

        'Add the CustomerAddRq aggregate to QBXMLMsgsRq aggregate
        Dim CustomerAddRqNode As MSXML2.IXMLDOMNode
		CustomerAddRqNode = xmlCustomerAdd.createElement("CustomerAddRq")
		QBXMLMsgsRqNode.appendChild(CustomerAddRqNode)
		
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Note - Starting requestID from 0 and incrementing it for each
		'request allows a developer to load the response into a QBFC and
		'parse it with QBFC
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		'Set the requestID attribute to 0
		Dim requestIDAttr As MSXML2.IXMLDOMAttribute
		requestIDAttr = xmlCustomerAdd.createAttribute("requestID")
		requestIDAttr.Text = "0"
        CustomerAddRqNode.attributes.setNamedItem(requestIDAttr)

        'Add the CustomerAdd aggregate to CustomerAddRq aggregate
        Dim CustomerAddNode As MSXML2.IXMLDOMNode
		CustomerAddNode = xmlCustomerAdd.createElement("CustomerAdd")
		CustomerAddRqNode.appendChild(CustomerAddNode)
		
		'Add the Name element to CustomerAdd aggregate
		Dim NameElement As MSXML2.IXMLDOMElement
		NameElement = xmlCustomerAdd.createElement("Name")
		NameElement.Text = strCustomerName
        CustomerAddNode.appendChild(NameElement)

        'Add the BillAddress aggregate to CustomerAdd aggregate
        Dim BillAddressNode As MSXML2.IXMLDOMNode
		BillAddressNode = xmlCustomerAdd.createElement("BillAddress")
		CustomerAddNode.appendChild(BillAddressNode)
		
		Dim Addr1Element As MSXML2.IXMLDOMElement
		If strAddr1 <> "" Then
			'Add the Addr1 element to BillAddress aggregate
			Addr1Element = xmlCustomerAdd.createElement("Addr1")
			Addr1Element.Text = strAddr1
            BillAddressNode.appendChild(Addr1Element)
        End If
		
		Dim Addr2Element As MSXML2.IXMLDOMElement
		If strAddr2 <> "" Then
			'Add the Addr2 element to BillAddress aggregate
			Addr2Element = xmlCustomerAdd.createElement("Addr2")
			Addr2Element.Text = strAddr2
            BillAddressNode.appendChild(Addr2Element)
        End If
		
		Dim Addr3Element As MSXML2.IXMLDOMElement
		If strAddr3 <> "" Then
			'Add the Addr3 element to BillAddress aggregate
			Addr3Element = xmlCustomerAdd.createElement("Addr3")
			Addr3Element.Text = strAddr3
            BillAddressNode.appendChild(Addr3Element)
        End If
		
		Dim Addr4Element As MSXML2.IXMLDOMElement
		If strAddr4 <> "" Then
			'Add the Addr4 element to BillAddress aggregate
			Addr4Element = xmlCustomerAdd.createElement("Addr4")
			Addr4Element.Text = strAddr4
            BillAddressNode.appendChild(Addr4Element)
        End If
		
		Dim CityElement As MSXML2.IXMLDOMElement
		If strCity <> "" Then
			'Add the City element to BillAddress aggregate
			CityElement = xmlCustomerAdd.createElement("City")
			CityElement.Text = strCity
            BillAddressNode.appendChild(CityElement)
        End If
		
		Dim StateProvinceElement As MSXML2.IXMLDOMElement
		If strStateProvince <> "" Then
			'Add the StateProvince element to BillAddress aggregate
			If strNationality = "US" Then
				StateProvinceElement = xmlCustomerAdd.createElement("State")
			Else
				StateProvinceElement = xmlCustomerAdd.createElement("Province")
			End If
			StateProvinceElement.Text = strStateProvince
            BillAddressNode.appendChild(StateProvinceElement)
        End If
		
		Dim PostalCodeElement As MSXML2.IXMLDOMElement
		If strPostalCode <> "" Then
			'Add the PostalCode element to BillAddress aggregate
			PostalCodeElement = xmlCustomerAdd.createElement("PostalCode")
			PostalCodeElement.Text = strPostalCode
            BillAddressNode.appendChild(PostalCodeElement)
        End If
		
		'We're adding the prolog using text strings
		If strNationality = "US" Then
			strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""2.0""?>" & rootElement.xml
		Else
			strXMLRequest = "<?xml version=""1.0"" ?>" & "<?qbxml version=""CA2.0""?>" & rootElement.xml
		End If
		
		On Error GoTo Errs
		
		strXMLResponse = qbXMLCOM.ProcessRequest(strTicket, strXMLRequest)
		
		MsgBox("Successfully added " & strCustomerName)
		AddCustomer = True
		Exit Function
		
Errs: 
		MsgBox("Error adding " & strCustomerName & " - HRESULT=" & Err.Number & " : " & Err.Description)
		AddCustomer = False
	End Function
	
	Public Sub EndQuickBooksSession()
		qbXMLCOM.EndSession(strTicket)
		qbXMLCOM.CloseConnection()
	End Sub
End Module