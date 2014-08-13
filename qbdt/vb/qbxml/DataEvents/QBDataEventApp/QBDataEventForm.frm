VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form QBDataEventForm 
   Caption         =   "QB Customer Change Tracker"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer CheckEventTimer 
      Interval        =   2000
      Left            =   120
      Top             =   840
   End
   Begin MSComctlLib.ListView CustomerList 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "QBDataEventForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This application is made up of 3 executables:
' QBDataEventApp        -- Provides the UI for this application (This app)
' QBDataEventSubscriber -- Subscribes the application to receive data events
'                          from QuickBooks.  This is the sort of application you
'                          would build as a Custom Action for an MSI installer to
'                          subscribe your application to events when you install your
'                          app and unsubscribe when your app is uninstalled.
' QBDataEventManager    -- Handles events from QuickBooks and supplies them to the main application
'                          via a second COM interface -- also has a UI to display the
'                          incoming event XML.
'
' Strictly speaking, we could probably do this in 2 apps (one to subscribe and one to
' do everything else) but to ensure that we process events from QuickBooks as quickly
' as possible and our main UI does not interfere with our event processing we divide it
' up, with the event manager here simply acting as a "broker" to the main application,
' serving events to it as requested.
'

' Variables to support access to QuickBooks
Private QBRP As QBXMLRP2Lib.RequestProcessor2
Private ticket As String

' Variables to keep track of the initial customer
' list we pull from QuickBooks
Private numCusts As Integer
Private IndexedCustArray() As CCustomer  'Customers already indexed by this app
Private IndexedCusts As Collection       'Used while we walk the results from QB
Private NewCusts As Collection           'Customers not yet indexed by this app
Private MaxIndex As Integer              'Maximum index stored in QuickBooks, note
                                         'that this is different from the number of
                                         'customers because we don't reuse an index if
                                         'a customer is merged or deleted
                                
                                
'Variables to support tracking events
'The QBEventHandler class is in the QBDataManager subtree of the DataEvent sample
'
Private EventManager As QBDataEventManager.QBEventHandler
Private Tracking As Boolean

'
' Most of the initial work happens as the form loads
'
Private Sub Form_Load()
    '
    'Don't start tracking events until we get the initial
    'list of customers filled in, but set up the event manager
    '
    Tracking = False
    Set EventManager = New QBDataEventManager.QBEventHandler
    
    '
    ' set up to talk with QuickBooks
    '
    Set QBRP = New QBXMLRP2Lib.RequestProcessor2
    QBRP.OpenConnection "", "IDN DataEventSample"
    ticket = QBRP.BeginSession("", qbFileOpenDoNotCare)
    
    SetupDataExt  'Make sure the data extension we need is defined
    
    '
    ' Set up the CustomerListView, note that we treat this control as a
    ' fake "database" for the purposes of this sample, the index of a
    ' customer in the list view will never change, and we use the index
    ' as the "key" to access the data from our ListView.
    '
    ' We take this approach so that you don't need anything other than the
    ' windows common controls to use the sample, no access, excel, etc. necessary!
    '
    CustomerList.View = lvwReport
    CustomerList.ColumnHeaders.Add , , "Name"
    CustomerList.ColumnHeaders.Add , , "ListID"
    CustomerList.ColumnHeaders.Add , , "Phone"
    CustomerList.ColumnHeaders.Add , , "eMail"
    CustomerList.ColumnHeaders.Add , , "Balance"
    
    '
    ' Get the customer list from QuickBooks
    '
    GetCustomers
    
    '
    ' Fill in the list view and add our data extension to each customer that
    ' needs it.
    FillListView
    
    '
    ' Now start watching for changes, see CheckEventTimer_Timer for
    ' processing of events.
    '
    Tracking = True
    EventManager.StartEventTracking
    
    '
    ' We don't hold onto a connection with QuickBooks because that can prevent
    ' the user from shutting down QuickBooks.
    '
    QBRP.EndSession ticket
    QBRP.CloseConnection
    Set QBRP = Nothing
    ticket = ""

End Sub

'
' Basically reverse everything we set up in the Load routine above
Private Sub Form_Unload(Cancel As Integer)
    Shutdown
End Sub

Public Sub Shutdown()
    '
    'tear down the connection to QuickBooks and drop the COM pointer
    '
    CheckEventTimer.Enabled = False
    Tracking = False
    If (Not QBRP Is Nothing) Then
        QBRP.EndSession ticket
        QBRP.CloseConnection
        Set QBRP = Nothing
        ticket = ""
    End If
    
    '
    ' Drop our COM pointers to the QBEventHandler
    '
    If (Not EventManager Is Nothing) Then
        EventManager.Shutdown
        Set EventManager = Nothing
    End If
End Sub

'
' Check for events we need to process
'
Private Sub CheckEventTimer_Timer()
    Dim eventXML As String
    If (Tracking) Then
        ' we are watching for events, so get one from the EventManager (if we have one)
        If (Not EventManager Is Nothing) Then
            eventXML = EventManager.GetEvent
        End If
        While (Not eventXML = "")
            ProcessEvent (eventXML)
            If (Not EventManager Is Nothing) Then
                eventXML = EventManager.GetEvent
            Else
                eventXML = ""
            End If
        Wend
    End If
End Sub

'
' Parse an event and decide what to do
'
Private Sub ProcessEvent(eventXML As String)
    Dim doc As DOMDocument40
    Set doc = New DOMDocument40
    If (doc.loadXML(eventXML)) Then
        '
        'FIRST check if QuickBooks is trying to shut down and follow suit
        '
        Dim UIEvents As IXMLDOMNodeList
        Set UIEvents = doc.getElementsByTagName("UIEvent")
        If (Not UIEvents.length = 0) Then
            'a company close event, let's shut ourselves down!
            Shutdown
            Unload Me
        Else
            ' NOW check for data events (since we aren't subscribing a
            ' menu extension that's the only other kind of event we can get
            '
            'Grab a connection to QuickBooks to process the event
            '
            Set QBRP = New QBXMLRP2Lib.RequestProcessor2
            QBRP.OpenConnection "", "IDN DataEventSample"
            ticket = QBRP.BeginSession("", qbFileOpenDoNotCare)

            '
            'a data event, something about our customers changed...
            '
            Dim ListEvents As IXMLDOMNodeList
            Set ListEvents = doc.getElementsByTagName("ListEvent")
            Dim ListEvent As IXMLDOMElement
            Set ListEvent = ListEvents.item(0) 'Should always be at least 1
            If (Not ListEvent Is Nothing) Then
                '
                ' Get the list operation that took place and the ListID involved
                '
                Dim op As String
                Dim id As String
                op = ListEvent.selectSingleNode("ListEventOperation").Text
                id = ListEvent.selectSingleNode("ListID").Text
                If (op = "Delete") Then
                    DelCust id
                ElseIf (op = "Add") Then
                    AddCust id
                ElseIf (op = "Modify") Then
                    ModCust id
                Else
                    'Ignore merge for now
                End If
            End If
            QBRP.EndSession ticket
            QBRP.CloseConnection
            Set QBRP = Nothing
            ticket = ""

        End If
    End If
End Sub

'
' Process a deleted customer event
'
Private Sub DelCust(id As String)
    ' This one is unique in that we can't go look it up in QuickBooks
    ' to find out our data extension that remembers its place in the
    ' customer list, so we have to find it the hard way in our data here
    Dim item As ListItem
    Set item = CustomerList.FindItem(id, 1)
    
    If (Not item Is Nothing) Then
        ' if we found the item, change it's color and show that it's been
        ' deleted.
        item.EnsureVisible
        item.ForeColor = vbYellow
        item.Bold = False
        item.Text = "#Deleted"
    End If
End Sub

'
' Process a customer add event
'
Private Sub AddCust(id As String)
    '
    ' First, get the particulars of the new customer from QuickBooks using
    ' the listID provided to us by the event
    Dim Cust As CCustomer
    Set Cust = GetCustomer(id)
    
    ' If we found the customer, add it to the list
    If (Not Cust Is Nothing) Then
        ' note that regardless of the alphabet, we just add the new customer to the
        ' end, color it blue to show that it is something that changed since we
        ' started
        MaxIndex = MaxIndex + 1
        With CustomerList.ListItems.Add(MaxIndex, Cust.ListID, Cust.name)
            .EnsureVisible
            .ListSubItems.Add , , Cust.ListID
            .ListSubItems.Add , , Cust.Phone
            .ListSubItems.Add , , Cust.Email
            .ListSubItems.Add , , Cust.Balance
            .ForeColor = vbBlue
            .Bold = True
        End With
        ' since it is a new customer, we need to data extend it to remember the index
        ' we just assigned it!
        '
        ' NOTE: This is *really* important -- by extending the data as a DIRECT result
        ' of an event, we put ourselves at risk of (by ourselves or in concert with
        ' other applications) getting into an infinite recursion.  Do this ONLY if
        ' absolutely necessary to your app and ONLY by taking the kind of precautions
        ' you see here - don't process any events that come in while your app is
        ' modifying data, etc.  See the ExtendData routine for more details.
        ExtendData Cust, MaxIndex
    End If
End Sub

'
' Process a customer modify event.  Note that all kinds of things cause this,
' a direct change to a customer from the UI, an application sending a data extension,
' a transaction (invoice, receive payment, etc.) that change a customer's balance, etc.
'
' Note, in particular, that we do NOT make any calls to QuickBooks to modify data in
' a record that we got a modify event for.  This ensures that we, at least, are not
' a part of an infinite event/response recursion because we only modify data as a
' result of an add event, which happens only once per record...
'
Private Sub ModCust(id As String)
    '
    ' First, get the particulars of the customer from QuickBooks
    '
    Dim Cust As CCustomer
    Set Cust = GetCustomer(id)
    
    '
    ' If we found the customer...
    If (Not Cust Is Nothing) Then
        '
        ' Get the customer from our "database" by the index we stored
        '
        Dim idx As Integer
        idx = Cust.index
        Dim item As ListItem
        Set item = CustomerList.ListItems.item(idx)
        
        '
        ' If we found the customer in our database, show the new info (note
        ' that the change might not be in any of the fields we're tracking, but
        ' that's ok.  And change the color to show that it has changed.
        '
        If (Not item Is Nothing) Then
            With item
                .Text = Cust.name
                .ListSubItems.item(1).Text = Cust.ListID
                .ListSubItems.item(2).Text = Cust.Phone
                .ListSubItems.item(3).Text = Cust.Email
                .ListSubItems.item(4).Text = Cust.Balance
                .ForeColor = vbRed
                .Bold = True
            End With
        End If
    End If
End Sub

'
' This is the all-important ExtendData routine.  It takes one of our customer
' records and the index of that customer in our "database" and saves the Index as
' a data extension.  Because this might happen as the direct result of an event,
' we must be very careful that we don't process the events that result from our
' actions here.  The approach we take in this sample is to tell the event manager
' to stop queuing events while QuickBooks processes our request.  This *does* create
' a small window for us to miss events from other apps (i.e. QBXMLRP2 processes
' requests one at a time, if we are waiting for QuickBooks to process our call to
' ProcessRequest and have turned off event queing while we were waiting, we might miss
' things that happened during that time.  The safety here is worth the tradeoff, and
' the problem could be solved by having "StopTracking" change the queue where events
' are tracked so we could do a check of that queue when we are done processing the
' response from QuickBooks to make sure we didn't miss anything that came from some
' other app or the QuickBooks UI...
'
Private Sub ExtendData(Cust As CCustomer, index As Integer)
    Cust.index = index
    
    ' Set up the DataExtAdd request
    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMElement
    Dim QBXMLMsgsRq As IXMLDOMElement
    Set QBXML = doc.createElement("QBXML")
    doc.appendChild QBXML
    Set QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
    Dim onError As IXMLDOMAttribute
    Set onError = doc.createAttribute("onError")
    onError.Value = "stopOnError"
    QBXMLMsgsRq.Attributes.setNamedItem onError
    QBXML.appendChild QBXMLMsgsRq
    Dim dataextdef As IXMLDOMElement
    Set dataextdef = QBXMLMsgsRq.appendChild(doc.createElement("DataExtAddRq"))
    Set dataextdef = dataextdef.appendChild(doc.createElement("DataExtAdd"))
    
    ' Add the OwnerID, DataExtName, and what kind of object we are extending...
    dataextdef.appendChild(doc.createElement("OwnerID")).Text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    dataextdef.appendChild(doc.createElement("DataExtName")).Text = "QBDEIndex"
    dataextdef.appendChild(doc.createElement("ListDataExtType")).Text = "Customer"
    
    ' Add the reference to the customer object we want to extend...
    Dim objref As IXMLDOMElement
    Set objref = dataextdef.appendChild(doc.createElement("ListObjRef"))
    objref.appendChild(doc.createElement("ListID")).Text = Cust.ListID
    
    ' Set the value of the extension to our "database" index
    dataextdef.appendChild(doc.createElement("DataExtValue")).Text = index
    
    ' Now form up the full request and add the prolog
    Dim request As String
    request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & doc.xml
    
    'save our request for review later, great for debugging problems by
    'using the qbXML Validator
    saveXMLStream "ExtendData", request
    
    '
    ' Now we are ready to process the request.  This will cause a modify event, but
    ' because we don't extend the data if we don't have to we won't have a looping
    ' problem.
    '
    ' To avoid coloring everything as modified, we're going to filter the events we
    ' get by keeping a list of ID's that we are modifying, when the event comes in
    ' for that list ID the ID will be removed from the filter list.
    '
    EventManager.AddFilter Cust.ListID
    Dim resp As String
    resp = QBRP.ProcessRequest(ticket, request)
    saveXMLStream "ExtendData-Response", resp
        
    '
    ' Not processing the response from ProcessRequest AT ALL, is generally a bad idea
    ' but since it is not intrinsic to the purpose of this sample (and we can survive
    ' if our index isn't saved) we do nothing with the response from QuickBooks here...
    '
End Sub

'
' This routine is used during the form load event to get the complete list
' of customers from QuickBooks at startup.  We will then track changes from there
' via events.
Private Sub GetCustomers()
    ' Initialize necessary collections
    Set IndexedCusts = New Collection  'Customers we've seen before
    Set NewCusts = New Collection      'Customers added since the last time this app
                                       'ran.
    numCusts = 0
    
    'Set up the Customer Query
    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMElement
    Dim QBXMLMsgsRq As IXMLDOMElement
    Set QBXML = doc.createElement("QBXML")
    doc.appendChild QBXML
    Set QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
    Dim onError As IXMLDOMAttribute
    Set onError = doc.createAttribute("onError")
    onError.Value = "stopOnError"
    QBXMLMsgsRq.Attributes.setNamedItem onError
    QBXML.appendChild QBXMLMsgsRq
    Dim query As IXMLDOMElement
    Set query = QBXMLMsgsRq.appendChild(doc.createElement("CustomerQueryRq"))
    '
    ' Add our OwnerID (which matches our subscriber ID but doesn't have to) so that
    ' we get our data extension back for any customers we've already looked at.
    '
    query.appendChild(doc.createElement("OwnerID")).Text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    
    '
    ' Add the prolog to our request and save the request to a file for debugging with
    ' the validator if needed.
    '
    Dim request As String
    request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & doc.xml
    saveXMLStream "GetCustomers", request
    
    ' Send the customer Query
    Dim response As String
    response = QBRP.ProcessRequest(ticket, request)
    saveXMLStream "GetCustomers-Response", response
    
    ' Process the response to get the customer information we care about
    Dim respDOM As New DOMDocument
    If (respDOM.loadXML(response)) Then
        Dim CustomerXML As IXMLDOMNodeList
        Set CustomerXML = respDOM.getElementsByTagName("CustomerRet")
        Dim CustomerRet As IXMLDOMElement
        Dim Cust As CCustomer
        MaxIndex = 0

        For Each CustomerRet In CustomerXML
            numCusts = numCusts + 1
            
            ' Create a new CCustomer object and fill it in from the data in the
            ' CustomerRet object
            Set Cust = New CCustomer
            'These first two are required fields, so they'll be there...
            Cust.name = CustomerRet.selectSingleNode("FullName").Text
            Cust.ListID = CustomerRet.selectSingleNode("ListID").Text
            
            ' The next few fields are optional, check if they exist before using
            If (Not CustomerRet.selectSingleNode("Phone") Is Nothing) Then
                Cust.Phone = CustomerRet.selectSingleNode("Phone").Text
            End If
            If (Not CustomerRet.selectSingleNode("Email") Is Nothing) Then
                Cust.Email = CustomerRet.selectSingleNode("Email").Text
            End If
            If (Not CustomerRet.selectSingleNode("Balance") Is Nothing) Then
                Cust.Balance = CustomerRet.selectSingleNode("Balance").Text
            End If
            
            ' Now check for any DataExtRet aggregates in the response, in particular
            ' we care about our QBDEIndex extension which indicates we've seen and
            ' processed the customer before so it already has a specific place in
            ' our CustomerListView.
            Dim DataExts As IXMLDOMNodeList
            Set DataExts = CustomerRet.selectNodes("DataExtRet")
            Dim DataExt As IXMLDOMElement
            For Each DataExt In DataExts
                If (DataExt.selectSingleNode("DataExtName").Text = "QBDEIndex") Then
                    Cust.index = DataExt.selectSingleNode("DataExtValue").Text
                End If
            Next
            
            ' Finally, put it in the appropriate collection depending on whether we've
            ' seen it before or not...
            If (Cust.index = "") Then
                NewCusts.Add Cust
            Else
                If (Cust.index > MaxIndex) Then
                    MaxIndex = Cust.index
                End If
                IndexedCusts.Add Cust, Cust.index
            End If
        Next
        
        'Having processed the complete list of customers returned from QuickBooks,
        'we now know the MaxIndex of all the customers previously processed, so we
        'can build an array with each indexed customer in its appropriate place
        'in the array, if a customer has been deleted since we processed it, etc. then
        'the array will have "holes" where the object at that position is "Nothing"
        'we'll take care of that when we fill in the CustomerListView to show the
        'missing record as #Deleted.
        ReDim IndexedCustArray(MaxIndex) As CCustomer
        For Each Cust In IndexedCusts
            Dim idx As Integer
            idx = Cust.index
            Set IndexedCustArray(idx) = Cust
        Next
    End If
End Sub

'
' The GetCustomers routine above takes care of getting the customer data from
' QuickBooks, this routine is responsible for displaying it in our "database"
' ListView common control
'
Private Sub FillListView()
    Dim Cust As CCustomer
    ' First walk the list of Customers that have been seen by this program
    ' before and display the records in the appropriate place
    For i = 1 To UBound(IndexedCustArray)
        Set Cust = IndexedCustArray(i)
        If (Cust Is Nothing) Then
            ' Customer must have been deleted since we first saw it, display it
            ' as deleted.
            CustomerList.ListItems.Add i, "Del" & i, "#Deleted"
        Else
            ' We have all the customer data, fill it into the ListView and all
            ' the columns thereof.
            With CustomerList.ListItems.Add(i, Cust.ListID, Cust.name)
                .ListSubItems.Add , , Cust.ListID
                .ListSubItems.Add , , Cust.Phone
                .ListSubItems.Add , , Cust.Email
                .ListSubItems.Add , , Cust.Balance
            End With
        End If
    Next
    
    ' Now we walk the list of customers that this program has NOT seen before and
    ' add them to the listView "database" saving the index we display each customer in
    ' by using a DataExtension.  Note that when this routine is called we generally
    ' have event tracking turned off because we haven't gotten to the point where we
    ' are watching for events yet.
    For Each Cust In NewCusts
        MaxIndex = MaxIndex + 1
        With CustomerList.ListItems.Add(MaxIndex, Cust.ListID, Cust.name)
            .ListSubItems.Add , , Cust.ListID
            .ListSubItems.Add , , Cust.Phone
            .ListSubItems.Add , , Cust.Email
            .ListSubItems.Add , , Cust.Balance
        End With
        ExtendData Cust, MaxIndex
    Next
End Sub

'
' This is a helper function for event processing.  When we get an event it has
' only the ListID or the TxnID of the object affected, this is to ensure that a
' rogue app doesn't get access to customer information simply by hijacking another
' application's subscription data (digitally signing your application also helps
' in this area.)  Since all we have is a ListID, we need to query for the rest of
' the details we need...
Private Function GetCustomer(ListID As String) As CCustomer
    'Set up the Customer Query
    Dim doc As New DOMDocument40
    Dim QBXML As IXMLDOMElement
    Dim QBXMLMsgsRq As IXMLDOMElement
    Set QBXML = doc.createElement("QBXML")
    doc.appendChild QBXML
    Set QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
    Dim onError As IXMLDOMAttribute
    Set onError = doc.createAttribute("onError")
    onError.Value = "stopOnError"
    QBXMLMsgsRq.Attributes.setNamedItem onError
    QBXML.appendChild QBXMLMsgsRq
    Dim query As IXMLDOMElement
    'Basic CustomerQuery...
    Set query = QBXMLMsgsRq.appendChild(doc.createElement("CustomerQueryRq"))
    'Add the ListID we need...
    query.appendChild(doc.createElement("ListID")).Text = ListID
    'And our OwnerID so that if we've seen this customer before we get our DataExt...
    query.appendChild(doc.createElement("OwnerID")).Text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    
    'Assemble the request and add the prolog.  As always, save the request so we can
    'debug with the validator later...
    Dim request As String
    request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & doc.xml
    saveXMLStream "GetCustomers", request
    
    ' Send the customer Query
    Dim response As String
    response = QBRP.ProcessRequest(ticket, request)
    saveXMLStream "GetCustomer-Response", response
    
    ' Process the response, this is very similar to what we did in GetCustomers
    ' except that we should have only one customer returned (since the ListID
    ' uniquely identifies one customer) and we don't need to maintain the collections,
    ' etc. that routine did.
    Dim respDOM As New DOMDocument
    If (respDOM.loadXML(response)) Then
        Dim CustomerXML As IXMLDOMNodeList
        Set CustomerXML = respDOM.getElementsByTagName("CustomerRet")
        Dim CustomerRet As IXMLDOMElement
        Dim Cust As CCustomer

        ' Only one really, but since getElementsByTagName returns a list it's
        ' easiest to just loop through it...
        For Each CustomerRet In CustomerXML
            Set Cust = New CCustomer
            Cust.name = CustomerRet.selectSingleNode("FullName").Text
            Cust.ListID = CustomerRet.selectSingleNode("ListID").Text
            If (Not CustomerRet.selectSingleNode("Phone") Is Nothing) Then
                Cust.Phone = CustomerRet.selectSingleNode("Phone").Text
            End If
            If (Not CustomerRet.selectSingleNode("Email") Is Nothing) Then
                Cust.Email = CustomerRet.selectSingleNode("Email").Text
            End If
            If (Not CustomerRet.selectSingleNode("Balance") Is Nothing) Then
                Cust.Balance = CustomerRet.selectSingleNode("Balance").Text
            End If
            Dim DataExts As IXMLDOMNodeList
            Set DataExts = CustomerRet.selectNodes("DataExtRet")
            Dim DataExt As IXMLDOMElement
            For Each DataExt In DataExts
                If (DataExt.selectSingleNode("DataExtName").Text = "QBDEIndex") Then
                    Cust.index = DataExt.selectSingleNode("DataExtValue").Text
                End If
            Next
        Next
        ' Return the customer we got
        Set GetCustomer = Cust
    End If
End Function

'
' Routine to ensure that the company file has a definition for our data extension
' We could Query for the definition and then create it if it does not exist, but
' instead we'll rely on the fact that we'll get an error that we can ignore if
' the definition exists.  So we just do a DataExtDefAdd request to define our
' private extension, an integer extension to the Customer object named QBDEIndex
'
Private Sub SetupDataExt()
    Dim doc As New DOMDocument
    Dim QBXML As IXMLDOMElement
    Dim QBXMLMsgsRq As IXMLDOMElement
    
    ' Set up the QBXML and MsgsRq "envelope" for the request
    Set QBXML = doc.createElement("QBXML")
    doc.appendChild QBXML
    Set QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
    Dim onError As IXMLDOMAttribute
    Set onError = doc.createAttribute("onError")
    onError.Value = "stopOnError"
    QBXMLMsgsRq.Attributes.setNamedItem onError
    QBXML.appendChild QBXMLMsgsRq
    
    'Set up the DataExtDefAdd request
    Dim dataextdef As IXMLDOMElement
    Set dataextdef = QBXMLMsgsRq.appendChild(doc.createElement("DataExtDefAddRq"))
    Set dataextdef = dataextdef.appendChild(doc.createElement("DataExtDefAdd"))
    dataextdef.appendChild(doc.createElement("OwnerID")).Text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
    dataextdef.appendChild(doc.createElement("DataExtName")).Text = "QBDEIndex"
    dataextdef.appendChild(doc.createElement("DataExtType")).Text = "INTTYPE"
    dataextdef.appendChild(doc.createElement("AssignToObject")).Text = "Customer"
    
    'Add the prolog and send the request, saving it to a file for debugging with the
    ' qbXML validator tool
    Dim request As String
    request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & vbCrLf & doc.xml
    saveXMLStream "SetupDataExt", request
    Dim resp As String
    resp = QBRP.ProcessRequest(ticket, request)
    saveXMLStream "SetupDataExt-Response", resp
    
    '
    ' Again, not parsing the response at all is hardly robust programming, but for
    ' the purposes of this sample it is sufficient, the mostly likely error would
    ' be that the definition already exists (because we've been run before) and
    ' we don't care, because that means the definition is there  Production software
    ' would obviously not do this...
    '
End Sub

'
' A useful little utility (note: uses the MS Scripting Runtime) for
' saving XML to a file, great for using the validator when we get an error.
'
Private Sub saveXMLStream(name As String, xml As String)
    Dim fname As String
    fname = App.Path & "\" & name & ".qbxml"
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Set ts = fso.CreateTextFile(fname, True)
    ts.Write (xml)
End Sub
