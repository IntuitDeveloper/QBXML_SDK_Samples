Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Text
Imports QBDataEventManager.QBEventHandler

Friend Class QBDataEventForm
    Inherits System.Windows.Forms.Form
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
    'UPGRADE_WARNING: Arrays in structure QBRP may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
    Private QBRP As QBXMLRP2Lib.RequestProcessor2
    Private ticket As String

    ' Variables to keep track of the initial customer
    ' list we pull from QuickBooks
    Private numCusts As Short
    Private IndexedCustArray() As CCustomer 'Customers already indexed by this app
    Private IndexedCusts As Collection 'Used while we walk the results from QB
    Private NewCusts As Collection 'Customers not yet indexed by this app
    Private MaxIndex As Short 'Maximum index stored in QuickBooks, note
    'that this is different from the number of
    'customers because we don't reuse an index if
    'a customer is merged or deleted


    'Variables to support tracking events
    'The QBEventHandler class is in the QBDataManager subtree of the DataEvent sample
    '
    Private EventManager As Object
    Private Tracking As Boolean

    '
    ' Most of the initial work happens as the form loads
    '
    Private Sub QBDataEventForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '
        'Don't start tracking events until we get the initial
        'list of customers filled in, but set up the event manager
        '
        Tracking = False
        Dim obj As Object = CreateObject("QBDataEventManager.QBEventHandler")
        EventManager = CreateObject("QBDataEventManager.QBEventHandler")

        '
        ' set up to talk with QuickBooks
        '
        QBRP = New QBXMLRP2Lib.RequestProcessor2
        QBRP.OpenConnection("", "IDN DataEventSample")
        ticket = QBRP.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)

        SetupDataExt() 'Make sure the data extension we need is defined

        '
        ' Set up the CustomerListView, note that we treat this control as a
        ' fake "database" for the purposes of this sample, the index of a
        ' customer in the list view will never change, and we use the index
        ' as the "key" to access the data from our ListView.
        '
        ' We take this approach so that you don't need anything other than the
        ' windows common controls to use the sample, no access, excel, etc. necessary!
        '
        CustomerList.View = System.Windows.Forms.View.Details
        CustomerList.Columns.Add("Name")
        CustomerList.Columns.Add("ListID")
        CustomerList.Columns.Add("Phone")
        CustomerList.Columns.Add("eMail")
        CustomerList.Columns.Add("Balance")

        '
        ' Get the customer list from QuickBooks
        '
        GetCustomers()

        '
        ' Fill in the list view and add our data extension to each customer that
        ' needs it.
        FillListView()

        '
        ' Now start watching for changes, see CheckEventTimer_Timer for
        ' processing of events.
        '
        Tracking = True
        'UPGRADE_WARNING: Couldn't resolve default property of object EventManager.StartEventTracking. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EventManager.StartEventTracking()

        '
        ' We don't hold onto a connection with QuickBooks because that can prevent
        ' the user from shutting down QuickBooks.
        '
        QBRP.EndSession(ticket)
        QBRP.CloseConnection()
        'UPGRADE_NOTE: Object QBRP may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        QBRP = Nothing
        ticket = ""

    End Sub

    '
    ' Basically reverse everything we set up in the Load routine above
    Private Sub QBDataEventForm_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Shutdown()
    End Sub

    Public Sub Shutdown()
        '
        'tear down the connection to QuickBooks and drop the COM pointer
        '
        CheckEventTimer.Enabled = False
        Tracking = False
        If (Not QBRP Is Nothing) Then
            QBRP.EndSession(ticket)
            QBRP.CloseConnection()
            'UPGRADE_NOTE: Object QBRP may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            QBRP = Nothing
            ticket = ""
        End If

        '
        ' Drop our COM pointers to the QBEventHandler
        '
        If (Not EventManager Is Nothing) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object EventManager.Shutdown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            EventManager.Shutdown()
            'UPGRADE_NOTE: Object EventManager may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            EventManager = Nothing
        End If
    End Sub

    '
    ' Check for events we need to process
    '
    Private Sub CheckEventTimer_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CheckEventTimer.Tick
        Dim eventXML As String
        If (Tracking) Then
            ' we are watching for events, so get one from the EventManager (if we have one)
            If (Not EventManager Is Nothing) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object EventManager.GetEvent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                eventXML = EventManager.GetEvent
            End If
            While (Not eventXML = "")
                ProcessEvent((eventXML))
                If (Not EventManager Is Nothing) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object EventManager.GetEvent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    eventXML = EventManager.GetEvent
                Else
                    eventXML = ""
                End If
            End While
        End If
    End Sub

    '
    ' Parse an event and decide what to do
    '
    Private Sub ProcessEvent(ByRef eventXML As String)
        Dim doc As MSXML2.DOMDocument60
        doc = New MSXML2.DOMDocument60
        Dim UIEvents As MSXML2.IXMLDOMNodeList
        Dim ListEvents As MSXML2.IXMLDOMNodeList
        Dim ListEvent As MSXML2.IXMLDOMElement
        Dim op As String
        Dim id As String
        If (doc.loadXML(eventXML)) Then
            '
            'FIRST check if QuickBooks is trying to shut down and follow suit
            '
            UIEvents = doc.getElementsByTagName("UIEvent")
            If (Not UIEvents.length = 0) Then
                'a company close event, let's shut ourselves down!
                Shutdown()
                Me.Close()
            Else
                ' NOW check for data events (since we aren't subscribing a
                ' menu extension that's the only other kind of event we can get
                '
                'Grab a connection to QuickBooks to process the event
                '
                QBRP = New QBXMLRP2Lib.RequestProcessor2
                QBRP.OpenConnection("", "IDN DataEventSample")
                ticket = QBRP.BeginSession("", QBXMLRP2Lib.QBFileMode.qbFileOpenDoNotCare)

                '
                'a data event, something about our customers changed...
                '
                ListEvents = doc.getElementsByTagName("ListEvent")
                ListEvent = ListEvents.item(0) 'Should always be at least 1
                If (Not ListEvent Is Nothing) Then
                    '
                    ' Get the list operation that took place and the ListID involved
                    '
                    op = ListEvent.selectSingleNode("ListEventOperation").text
                    id = ListEvent.selectSingleNode("ListID").text
                    If (op = "Delete") Then
                        DelCust(id)
                    ElseIf (op = "Add") Then
                        AddCust(id)
                    ElseIf (op = "Modify") Then
                        ModCust(id)
                    Else
                        'Ignore merge for now
                    End If
                End If
                QBRP.EndSession(ticket)
                QBRP.CloseConnection()
                'UPGRADE_NOTE: Object QBRP may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                QBRP = Nothing
                ticket = ""

            End If
        End If
    End Sub

    '
    ' Process a deleted customer event
    '
    Private Sub DelCust(ByRef id As String)
        ' This one is unique in that we can't go look it up in QuickBooks
        ' to find out our data extension that remembers its place in the
        ' customer list, so we have to find it the hard way in our data here
        Dim item As System.Windows.Forms.ListViewItem
        item = CustomerList.FindItemWithText(id, True, 0)

        If (Not item Is Nothing) Then
            ' if we found the item, change it's color and show that it's been
            ' deleted.
            'UPGRADE_WARNING: MSComctlLib.ListItem method item.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            item.EnsureVisible()
            item.ForeColor = System.Drawing.Color.Yellow
            item.Font = New Font(CustomerList.Font, FontStyle.Regular) 'VB6.FontChangeBold(item.Font, False)
            item.Text = "#Deleted"
        End If
    End Sub

    '
    ' Process a customer add event
    '
    Private Sub AddCust(ByRef id As String)
        '
        ' First, get the particulars of the new customer from QuickBooks using
        ' the listID provided to us by the event
        Dim Cust As CCustomer
        Cust = GetCustomer(id)

        ' If we found the customer, add it to the list
        If (Not Cust Is Nothing) Then
            ' note that regardless of the alphabet, we just add the new customer to the
            ' end, color it blue to show that it is something that changed since we
            ' started
            'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            With CustomerList.Items.Insert(MaxIndex, Cust.name, Cust.name, "")
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: MSComctlLib.IListItem method CustomerList.ListItems.Add.EnsureVisible has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
                .EnsureVisible()
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.ListID)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.Phone)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.Email)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.Balance)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .ForeColor = System.Drawing.Color.Blue
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .Font = New Font(CustomerList.Font, FontStyle.Bold) 'VB6.FontChangeBold(.Font, True)
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
            ExtendData(Cust, MaxIndex)
            MaxIndex = MaxIndex + 1
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
    Private Sub ModCust(ByRef id As String)
        '
        ' First, get the particulars of the customer from QuickBooks
        '
        Dim Cust As CCustomer
        Cust = GetCustomer(id)

        '
        ' If we found the customer...
        Dim idx As Short
        Dim item As System.Windows.Forms.ListViewItem
        If (Not Cust Is Nothing) Then
            '
            ' Get the customer from our "database" by the index we stored
            '
            idx = CShort(Cust.index)
            'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            item = CustomerList.Items.Item(idx)

            '
            ' If we found the customer in our database, show the new info (note
            ' that the change might not be in any of the fields we're tracking, but
            ' that's ok.  And change the color to show that it has changed.
            '
            If (Not item Is Nothing) Then
                With item
                    .Text = Cust.name
                    'UPGRADE_WARNING: Lower bound of collection item.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    .SubItems.Item(1).Text = Cust.ListID
                    'UPGRADE_WARNING: Lower bound of collection item.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    .SubItems.Item(2).Text = Cust.Phone
                    'UPGRADE_WARNING: Lower bound of collection item.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    .SubItems.Item(3).Text = Cust.Email
                    'UPGRADE_WARNING: Lower bound of collection item.ListSubItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    .SubItems.Item(4).Text = Cust.Balance
                    .ForeColor = System.Drawing.Color.Red
                    .Font = New Font(CustomerList.Font, FontStyle.Bold) 'VB6.FontChangeBold(.Font, True)
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
    Private Sub ExtendData(ByRef Cust As CCustomer, ByRef index As Short)
        Cust.index = CStr(index)

        ' Set up the DataExtAdd request
        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMElement
        Dim QBXMLMsgsRq As MSXML2.IXMLDOMElement
        QBXML = doc.createElement("QBXML")
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        doc.appendChild(QBXML)
        QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
        Dim onError As MSXML2.IXMLDOMAttribute
        onError = doc.createAttribute("onError")
        onError.value = "stopOnError"
        'UPGRADE_WARNING: Couldn't resolve default property of object onError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXMLMsgsRq.attributes.setNamedItem(onError)
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXMLMsgsRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXML.appendChild(QBXMLMsgsRq)
        Dim dataextdef As MSXML2.IXMLDOMElement
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef = QBXMLMsgsRq.appendChild(doc.createElement("DataExtAddRq"))
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef = dataextdef.appendChild(doc.createElement("DataExtAdd"))

        ' Add the OwnerID, DataExtName, and what kind of object we are extending...
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("OwnerID")).text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("DataExtName")).text = "QBDEIndex"
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("ListDataExtType")).text = "Customer"

        ' Add the reference to the customer object we want to extend...
        Dim objref As MSXML2.IXMLDOMElement
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objref = dataextdef.appendChild(doc.createElement("ListObjRef"))
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objref.appendChild(doc.createElement("ListID")).text = Cust.ListID

        ' Set the value of the extension to our "database" index
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("DataExtValue")).text = CStr(index)

        ' Now form up the full request and add the prolog
        Dim request As String
        request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & doc.xml

        'save our request for review later, great for debugging problems by
        'using the qbXML Validator
        saveXMLStream("ExtendData", request)

        '
        ' Now we are ready to process the request.  This will cause a modify event, but
        ' because we don't extend the data if we don't have to we won't have a looping
        ' problem.
        '
        ' To avoid coloring everything as modified, we're going to filter the events we
        ' get by keeping a list of ID's that we are modifying, when the event comes in
        ' for that list ID the ID will be removed from the filter list.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object EventManager.AddFilter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        EventManager.AddFilter(Cust.ListID)
        Dim resp As String
        resp = QBRP.ProcessRequest(ticket, request)
        saveXMLStream("ExtendData-Response", resp)

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
        IndexedCusts = New Collection 'Customers we've seen before
        NewCusts = New Collection 'Customers added since the last time this app
        'ran.
        numCusts = 0

        'Set up the Customer Query
        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMElement
        Dim QBXMLMsgsRq As MSXML2.IXMLDOMElement
        QBXML = doc.createElement("QBXML")
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        doc.appendChild(QBXML)
        QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
        Dim onError As MSXML2.IXMLDOMAttribute
        onError = doc.createAttribute("onError")
        onError.value = "stopOnError"
        'UPGRADE_WARNING: Couldn't resolve default property of object onError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXMLMsgsRq.attributes.setNamedItem(onError)
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXMLMsgsRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXML.appendChild(QBXMLMsgsRq)
        Dim query As MSXML2.IXMLDOMElement
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        query = QBXMLMsgsRq.appendChild(doc.createElement("CustomerQueryRq"))
        '
        ' Add our OwnerID (which matches our subscriber ID but doesn't have to) so that
        ' we get our data extension back for any customers we've already looked at.
        '
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        query.appendChild(doc.createElement("OwnerID")).text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"

        '
        ' Add the prolog to our request and save the request to a file for debugging with
        ' the validator if needed.
        '
        Dim request As String
        request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & doc.xml
        saveXMLStream("GetCustomers", request)

        ' Send the customer Query
        Dim response As String
        response = QBRP.ProcessRequest(ticket, request)
        saveXMLStream("GetCustomers-Response", response)

        ' Process the response to get the customer information we care about
        Dim respDOM As New MSXML2.DOMDocument60
        Dim CustomerXML As MSXML2.IXMLDOMNodeList
        Dim CustomerRet As MSXML2.IXMLDOMElement
        Dim Cust As CCustomer
        Dim DataExts As MSXML2.IXMLDOMNodeList
        Dim DataExt As MSXML2.IXMLDOMElement
        Dim idx As Short
        If (respDOM.loadXML(response)) Then
            CustomerXML = respDOM.getElementsByTagName("CustomerRet")
            MaxIndex = 0

            For Each CustomerRet In CustomerXML
                ' Create a new CCustomer object and fill it in from the data in the
                ' CustomerRet object
                Cust = New CCustomer
                'These first two are required fields, so they'll be there...
                Cust.name = CustomerRet.selectSingleNode("FullName").text
                Cust.ListID = CustomerRet.selectSingleNode("ListID").text

                ' The next few fields are optional, check if they exist before using
                If (Not CustomerRet.selectSingleNode("Phone") Is Nothing) Then
                    Cust.Phone = CustomerRet.selectSingleNode("Phone").text
                End If
                If (Not CustomerRet.selectSingleNode("Email") Is Nothing) Then
                    Cust.Email = CustomerRet.selectSingleNode("Email").text
                End If
                If (Not CustomerRet.selectSingleNode("Balance") Is Nothing) Then
                    Cust.Balance = CustomerRet.selectSingleNode("Balance").text
                End If

                ' Now check for any DataExtRet aggregates in the response, in particular
                ' we care about our QBDEIndex extension which indicates we've seen and
                ' processed the customer before so it already has a specific place in
                ' our CustomerListView.
                DataExts = CustomerRet.selectNodes("DataExtRet")
                For Each DataExt In DataExts
                    If (DataExt.selectSingleNode("DataExtName").text = "QBDEIndex") Then
                        Cust.index = DataExt.selectSingleNode("DataExtValue").text
                    End If
                Next DataExt

                ' Finally, put it in the appropriate collection depending on whether we've
                ' seen it before or not...
                If (Cust.index = "") Then
                    NewCusts.Add(Cust)
                Else
                    If (CDbl(Cust.index) > MaxIndex) Then
                        MaxIndex = CShort(Cust.index)
                    End If
                    IndexedCusts.Add(Cust, Cust.index)
                End If
                numCusts = numCusts + 1
            Next CustomerRet

            'Having processed the complete list of customers returned from QuickBooks,
            'we now know the MaxIndex of all the customers previously processed, so we
            'can build an array with each indexed customer in its appropriate place
            'in the array, if a customer has been deleted since we processed it, etc. then
            'the array will have "holes" where the object at that position is "Nothing"
            'we'll take care of that when we fill in the CustomerListView to show the
            'missing record as #Deleted.
            If IndexedCusts.Count > 0 Then
                ReDim IndexedCustArray(MaxIndex)
            End If

            For Each Cust In IndexedCusts
                    idx = CShort(Cust.index)
                    IndexedCustArray(idx) = Cust
                Next Cust
            End If
    End Sub

    '
    ' The GetCustomers routine above takes care of getting the customer data from
    ' QuickBooks, this routine is responsible for displaying it in our "database"
    ' ListView common control
    '
    Private Sub FillListView()
        Dim i As Object
        Dim Cust As CCustomer
        ' First walk the list of Customers that have been seen by this program
        ' before and display the records in the appropriate place
        If IndexedCustArray IsNot Nothing Then
            For i = 0 To IndexedCustArray.Length - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Cust = IndexedCustArray(i)
                If (Cust Is Nothing) Then
                    ' Customer must have been deleted since we first saw it, display it
                    ' as deleted.
                    'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_ISSUE: MSComctlLib.ListItems method CustomerList.ListItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    CustomerList.Items.Add(i, "Del" & i, "#Deleted")
                Else
                    ' We have all the customer data, fill it into the ListView and all
                    ' the columns thereof.
                    'UPGRADE_ISSUE: MSComctlLib.ListItems method CustomerList.ListItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    With CustomerList.Items.Add(i, Cust.name, Cust.name)
                        'UPGRADE_ISSUE: MSComctlLib.ListItems method CustomerList.ListItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        .SubItems.Add(Cust.ListID)
                        'UPGRADE_ISSUE: MSComctlLib.ListItems method CustomerList.ListItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        .SubItems.Add(Cust.Phone)
                        'UPGRADE_ISSUE: MSComctlLib.ListItems method CustomerList.ListItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        .SubItems.Add(Cust.Email)
                        'UPGRADE_ISSUE: MSComctlLib.ListItems method CustomerList.ListItems.Add was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                        .SubItems.Add(Cust.Balance)
                    End With
                End If
            Next
        End If

        ' Now we walk the list of customers that this program has NOT seen before and
        ' add them to the listView "database" saving the index we display each customer in
        ' by using a DataExtension.  Note that when this routine is called we generally
        ' have event tracking turned off because we haven't gotten to the point where we
        ' are watching for events yet.
        For Each Cust In NewCusts            'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            With CustomerList.Items.Insert(MaxIndex, Cust.name, Cust.name, "")
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.ListID)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.Phone)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.Email)
                'UPGRADE_WARNING: Lower bound of collection CustomerList.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                .SubItems.Add(Cust.Balance)
            End With
            ExtendData(Cust, MaxIndex)
            MaxIndex = MaxIndex + 1
        Next Cust
    End Sub

    '
    ' This is a helper function for event processing.  When we get an event it has
    ' only the ListID or the TxnID of the object affected, this is to ensure that a
    ' rogue app doesn't get access to customer information simply by hijacking another
    ' application's subscription data (digitally signing your application also helps
    ' in this area.)  Since all we have is a ListID, we need to query for the rest of
    ' the details we need...
    Private Function GetCustomer(ByRef ListID As String) As CCustomer
        'Set up the Customer Query
        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMElement
        Dim QBXMLMsgsRq As MSXML2.IXMLDOMElement
        QBXML = doc.createElement("QBXML")
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        doc.appendChild(QBXML)
        QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
        Dim onError As MSXML2.IXMLDOMAttribute
        onError = doc.createAttribute("onError")
        onError.value = "stopOnError"
        'UPGRADE_WARNING: Couldn't resolve default property of object onError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXMLMsgsRq.attributes.setNamedItem(onError)
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXMLMsgsRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXML.appendChild(QBXMLMsgsRq)
        Dim query As MSXML2.IXMLDOMElement
        'Basic CustomerQuery...
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        query = QBXMLMsgsRq.appendChild(doc.createElement("CustomerQueryRq"))
        'Add the ListID we need...
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        query.appendChild(doc.createElement("ListID")).text = ListID
        'And our OwnerID so that if we've seen this customer before we get our DataExt...
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        query.appendChild(doc.createElement("OwnerID")).text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"

        'Assemble the request and add the prolog.  As always, save the request so we can
        'debug with the validator later...
        Dim request As String
        request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & doc.xml
        saveXMLStream("GetCustomers", request)

        ' Send the customer Query
        Dim response As String
        response = QBRP.ProcessRequest(ticket, request)
        saveXMLStream("GetCustomer-Response", response)

        ' Process the response, this is very similar to what we did in GetCustomers
        ' except that we should have only one customer returned (since the ListID
        ' uniquely identifies one customer) and we don't need to maintain the collections,
        ' etc. that routine did.
        Dim respDOM As New MSXML2.DOMDocument60
        Dim CustomerXML As MSXML2.IXMLDOMNodeList
        Dim CustomerRet As MSXML2.IXMLDOMElement
        Dim Cust As CCustomer
        Dim DataExts As MSXML2.IXMLDOMNodeList
        Dim DataExt As MSXML2.IXMLDOMElement
        If (respDOM.loadXML(response)) Then
            CustomerXML = respDOM.getElementsByTagName("CustomerRet")

            ' Only one really, but since getElementsByTagName returns a list it's
            ' easiest to just loop through it...
            For Each CustomerRet In CustomerXML
                Cust = New CCustomer
                Cust.name = CustomerRet.selectSingleNode("FullName").text
                Cust.ListID = CustomerRet.selectSingleNode("ListID").text
                If (Not CustomerRet.selectSingleNode("Phone") Is Nothing) Then
                    Cust.Phone = CustomerRet.selectSingleNode("Phone").text
                End If
                If (Not CustomerRet.selectSingleNode("Email") Is Nothing) Then
                    Cust.Email = CustomerRet.selectSingleNode("Email").text
                End If
                If (Not CustomerRet.selectSingleNode("Balance") Is Nothing) Then
                    Cust.Balance = CustomerRet.selectSingleNode("Balance").text
                End If
                DataExts = CustomerRet.selectNodes("DataExtRet")
                For Each DataExt In DataExts
                    If (DataExt.selectSingleNode("DataExtName").text = "QBDEIndex") Then
                        Cust.index = DataExt.selectSingleNode("DataExtValue").text
                    End If
                Next DataExt
            Next CustomerRet
            ' Return the customer we got
            GetCustomer = Cust
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
        Dim doc As New MSXML2.DOMDocument60
        Dim QBXML As MSXML2.IXMLDOMElement
        Dim QBXMLMsgsRq As MSXML2.IXMLDOMElement

        ' Set up the QBXML and MsgsRq "envelope" for the request
        QBXML = doc.createElement("QBXML")
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXML. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        doc.appendChild(QBXML)
        QBXMLMsgsRq = doc.createElement("QBXMLMsgsRq")
        Dim onError As MSXML2.IXMLDOMAttribute
        onError = doc.createAttribute("onError")
        onError.value = "stopOnError"
        'UPGRADE_WARNING: Couldn't resolve default property of object onError. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXMLMsgsRq.attributes.setNamedItem(onError)
        'UPGRADE_WARNING: Couldn't resolve default property of object QBXMLMsgsRq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        QBXML.appendChild(QBXMLMsgsRq)

        'Set up the DataExtDefAdd request
        Dim dataextdef As MSXML2.IXMLDOMElement
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef = QBXMLMsgsRq.appendChild(doc.createElement("DataExtDefAddRq"))
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef = dataextdef.appendChild(doc.createElement("DataExtDefAdd"))
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("OwnerID")).text = "{2B6C9DB4-EBE2-45E7-A14F-4E1C49C965F7}"
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("DataExtName")).text = "QBDEIndex"
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("DataExtType")).text = "INTTYPE"
        'UPGRADE_WARNING: Couldn't resolve default property of object doc.createElement(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        dataextdef.appendChild(doc.createElement("AssignToObject")).text = "Customer"

        'Add the prolog and send the request, saving it to a file for debugging with the
        ' qbXML validator tool
        Dim request As String
        request = "<?xml version=""1.0""?>" & vbCrLf & "<?qbxml version=""3.0""?>" & vbCrLf & doc.xml
        saveXMLStream("SetupDataExt", request)
        Dim resp As String
        resp = QBRP.ProcessRequest(ticket, request)
        saveXMLStream("SetupDataExt-Response", resp)

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
    'UPGRADE_NOTE: name was upgraded to name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub saveXMLStream(ByRef name_Renamed As String, ByRef xml As String)
        Dim fname As String
        fname = My.Application.Info.DirectoryPath & "\" & name_Renamed & ".qbxml"
        ' Create or overwrite the file.
        Dim fs As FileStream = File.Create(fname) ' Add text to the file.
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(xml)
        fs.Write(info, 0, info.Length)
        fs.Close()
    End Sub
End Class