Attribute VB_Name = "ItemInventoryModule"
'----------------------------------------------------------
' Module: ItemInventoryModule
'
' Description: This module contains the code which creates QBFC
'              messages, exchanges them with QuickBooks, interprets
'              the responses and loads information into form objects.
'
' Routines: OpenConnectionBeginSession
'             Opens a connection and begins a sesson with the
'             currently open company file.  If a company isn't open,
'             the routine will display a message and then exit the
'             program.
'
'           EndSessionCloseConnection
'             Calls EndSession and CloseConnection if the boolean
'             booSessionBegun is true.
'
'           FillItemInventoryListBox
'             Builds an ItemInventoryQuery request for all items.
'             Fills the Item Inventory List Box with the FullName/ListID
'             object for each Item.  The Item's FullName is displayed in
'             list box.  Parent Combo box is also filled with this list.
'
'           QueryItemInventoryFields
'             Builds an ItemInventoryQuery for a specific ItemInventory
'             by specifying the ListID.  Then the DisplayItemInventoryFields
'             function is called to plug the field values into the form.
'
'           ModifyItemInventoryFields
'             Builds an ItemInventoryMod for a specific ItemInventory
'             by specifying the ListID.  All possible field values to be
'             modified are filled in with the values from the form.
'             Then the DisplayItemInventoryFields function is called
'             to plug the field values into the form.
'
'           DisplayItemInventoryFields
'             Fill the fields in the main form with the values from
'             an IItemInventoryRetList object.
'
'           FillAssetAccountList
'             Builds an AccountQuery request for all AssetAccount items.
'             Fills the AssetAccount Combo Box with the FullName/ListID
'             object for each Item.  The Item's Name is displayed in
'             combo box.
'
'           FillCOGSAccountList
'             Builds an AccountQuery request for all COGSAccount items.
'             Fills the COGSAccount Combo Box with the FullName/ListID
'             object for each Item.  The Item's Name is displayed in
'             combo box.
'
'           FillPrefVendorList
'             Builds a VendorQuery request for all Vendor names.
'             Fills the PrefVendor Combo Box with the FullName/ListID
'             object for each Item.  The Item's Name is displayed in
'             combo box.
'
'           FillSalesTaxCodeList
'             Builds a SalesTaxCodeQuery request for all names.
'             Fills the SalesTaxCode Combo Box with the FullName/ListID
'             object for each Item.  The Item's Name is displayed in
'             combo box.
'
'           FindFullNameInComboBox
'             Loops thru the item objects in the combo box and returns
'             index of the item with the name.
'
'           FoundInComboBox
'             Loops thru the item objects in the combo box and returns
'             true if one of the objects has the specified name.
'
' Copyright © 2002-2013 Intuit Inc. All rights reserved.
' Use is subject to the terms specified at:
'      http://developer.intuit.com/legal/devsite_tos.html
'
' Updated to QBFC 12.0 and fixed setting max QBXML version: 09/2012
'----------------------------------------------------------
    Dim booSessionBegun As Boolean
    Dim booConnectionOpened As Boolean

    Const MAX_RETURNED = 20

    Public itemCollection As Collection
    Dim salesTaxCodeCollection As Collection
    Dim COGSCollection As Collection
    Dim vendorCollection As Collection
    Dim assetAccountCollection As Collection
    
    'Module objects
    Dim QBSessionManager As QBSessionManager
    Dim msgSetRequest As IMsgSetRequest
    

    Public Sub OpenConnectionBeginSession()

        On Error GoTo Errs

        If (booSessionBegun) Then
            Exit Sub
        End If
        
        ' create the new QBSessionManager object
        If (QBSessionManager Is Nothing) Then
            Set QBSessionManager = New QBSessionManager
        End If

        ' open the connection to QuickBooks
        If (Not booConnectionOpened) Then
            QBSessionManager.OpenConnection "", "IDN Item Inventory Modify Sample - QBFC"
            booConnectionOpened = True
        End If

        ' Begin a session with QuickBooks
        QBSessionManager.BeginSession "", ENOpenMode.omDontCare
        booSessionBegun = True

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
        Dim supportedVersion As String
        supportedVersion = Val(QBFCLatestVersion(QBSessionManager))
        If (supportedVersion >= 6#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 6, 0)
        ElseIf (supportedVersion >= 5#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 5, 0)
        ElseIf (supportedVersion >= 4#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 4, 0)
        ElseIf (supportedVersion >= 3#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 3, 0)
        ElseIf (supportedVersion >= 2#) Then
            booSupports2dot0 = True
            Set msgSetRequest = QBSessionManager.CreateMsgSetRequest("US", 2, 0)
        End If
        
        If Not booSupports2dot0 Then
            MsgBox "This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0"
            If booSessionBegun Then
                QBSessionManager.EndSession
                booSessionBegun = False
            End If
            If booConnectionOpened Then
                QBSessionManager.CloseConnection
                booConnectionOpened = False
            End If
            End
        End If

        Exit Sub

Errs:
        If Err.Number = &H80040416 Then
            MsgBox ("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
            QBSessionManager.CloseConnection
            booConnectionOpened = False
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox ("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
            QBSessionManager.CloseConnection
            booConnectionOpened = False
            End
        Else
            MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, vbOKOnly, "Error in OpenConnectionBeginSession"

            If booSessionBegun Then
                QBSessionManager.EndSession
                booSessionBegun = False
            End If

            If booConnectionOpened Then
                QBSessionManager.CloseConnection
                booConnectionOpened = False
            End If
            End
        End If
    End Sub

    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs
        If booSessionBegun Then
            QBSessionManager.EndSession
            booSessionBegun = False
        End If
        If booConnectionOpened Then
            QBSessionManager.CloseConnection
            booConnectionOpened = False
        End If
        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in EndSessionCloseConnection"
    End Sub

    Public Sub FillItemInventoryListBox(ByRef itemInventoryList As ListBox, _
                        ByRef ParentComboBox As comboBox)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If
        
        '*** first query QB for the all of Item Inventory List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryQuery As IItemInventoryQuery
        Set itemInventoryQuery = msgSetRequest.AppendItemInventoryQueryRq()

        ' lets get all the Item Inventory, active and inactive.
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ActiveStatus.SetValue (ENActiveStatus.asAll)

        ' we are going to set the number of entries returned to limit the size of the return structure
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.MaxReturned.SetValue (MAX_RETURNED)

        Dim bDone As Boolean
        bDone = False
        Dim firstFullName As String
        firstFullName = "!"

        'Clear the list box and collections
        itemInventoryList.Clear
        ParentComboBox.Clear
        If (itemCollection Is Nothing) Then
            Set itemCollection = New Collection
        End If
        Do While (itemCollection.count > 0)
            itemCollection.Remove (1)
        Loop

        Do While (Not bDone)

            ' start looking for item next in the list
            ' we will have one overlap
            itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.FromName.SetValue (firstFullName)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

            '*** Now go thru the response from QB and add the Item Inventory names ***
            '*** to the list box ***

            ' check to make sure we have objects to access first
            ' and that there are responses in the list
            If (msgSetResponse Is Nothing) Then
                Exit Sub
            End If
            If (msgSetResponse.responseList Is Nothing) Then
                Exit Sub
            End If
            If (msgSetResponse.responseList.count <= 0) Then
                Exit Sub
            End If

            ' Start parsing the response list
            Dim responseList As IResponseList
            Set responseList = msgSetResponse.responseList

            ' go thru each response and process the response.
            ' this example will only have one response in the list
            ' so we will look at index=0
            Dim response As IResponse
            Set response = responseList.GetAt(0)
            If (Not response Is Nothing) Then
                If response.StatusCode <> "0" Then
                    'If the status is bad, report it to the user
                    MsgBox ("FillItemInventoryListBox unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                    Exit Sub
                End If
            End If

            ' first make sure we have a response object to handle
            If (response Is Nothing) Then
                Exit Sub
            End If
            If (response.Type Is Nothing) Or _
                (response.Detail Is Nothing) Then
                Exit Sub
            End If
            If (response.Detail.Type Is Nothing) Then
                Exit Sub
            End If

            ' make sure we are processing the ItemInventoryQueryRs and
            ' the ItemInventoryRetList responses in this response list
            Dim itemInventoryRetList As IItemInventoryRetList
            Dim responseType As ENResponseType
            Dim responseDetailType As ENObjectType
            responseType = response.Type.GetValue()
            responseDetailType = response.Detail.Type.GetValue()
            If (responseType = ENResponseType.rtItemInventoryQueryRs) And _
                (responseDetailType = ENObjectType.otItemInventoryRetList) Then
                ' save the response detail in the appropriate object type
                ' since we have first verified the type of the response object
                Set itemInventoryRetList = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                Exit Sub
            End If

            'Parse the query response and add the Item to the Item Inventory list box
            Dim count As Integer
            Dim index As Integer
            Dim itemInventoryRet As IItemInventoryRet
            Dim itemInfo As FullNameListIDClass
            count = itemInventoryRetList.count

            ' we are done with the itemInventoryQueries if we have not received the MaxReturned
            If (count < MAX_RETURNED) Then
                bDone = True
            End If

            For index = 0 To count - 1
                Set itemInventoryRet = itemInventoryRetList.GetAt(index)
                If (Not itemInventoryRet Is Nothing) Then
                    If (Not itemInventoryRet.FullName Is Nothing) And _
                        (Not itemInventoryRet.ListID Is Nothing) Then
                        ' we are saving the FullName and ListID pairs for each Item Inventory
                        ' this is good practice since the FullName can change but the
                        ' ListID will not change for an Item Inventory.
                        If (index <> 0) Or (Not FoundListIDInCollection(itemCollection, itemInventoryRet.ListID.GetValue())) Then
                            ' if the item is not in the combo box, then it is also not in the list box
                            Set itemInfo = New FullNameListIDClass
                            itemInfo.FullName = itemInventoryRet.FullName.GetValue()
                            itemInfo.ListID = itemInventoryRet.ListID.GetValue()
                            itemInventoryList.AddItem itemInfo.FullName
                            ParentComboBox.AddItem itemInfo.FullName
                            itemCollection.Add itemInfo, itemInventoryRet.ListID.GetValue()
                        End If
                        firstFullName = itemInventoryRet.FullName.GetValue()
                    End If
                End If
            Next
        Loop


        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillItemInventoryListBox"
    End Sub

    Public Sub QueryItemInventoryFields(ByRef itemIndex As Integer, _
                        ByRef MainForm As MainForm)
        On Error GoTo Errs

        Dim itemInfo As FullNameListIDClass
        Set itemInfo = itemCollection.Item(itemIndex + 1)
        If (itemInfo Is Nothing) Then
            Exit Sub
        End If

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If
        
        '*** first query QB for this specific Item Inventory ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryQuery As IItemInventoryQuery
        Set itemInventoryQuery = msgSetRequest.AppendItemInventoryQueryRq()

        ' specify the ListID of the item we want returned
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListIDList.Add (itemInfo.ListID)


        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the Item Inventory fields ***
        '*** to the controls of the dialog box ***

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("QueryItemInventoryFields unexpexcted Error - " & vbCrLf & "Status Code = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the ItemInventoryQueryRs and
        ' the ItemInventoryRetList responses in this response list
        Dim itemInventoryRetList As IItemInventoryRetList
        Dim itemInventoryRet As IItemInventoryRet
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtItemInventoryQueryRs) And _
            (responseDetailType = ENObjectType.otItemInventoryRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set itemInventoryRetList = response.Detail
            If (itemInventoryRetList.count >= 0) Then
                Set itemInventoryRet = itemInventoryRetList.GetAt(0)
            End If
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        DisplayItemInventoryFields itemInventoryRet, MainForm

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in QueryItemInventoryListBox"
    End Sub

    Public Function ModifyItemInventoryFields(ByRef itemInfoClass As FullNameListIDClass, _
                        ByRef MainForm As MainForm) As String
        On Error GoTo Errs

        Dim objectInfo As FullNameListIDClass
        Dim listIndex As Integer
        
        ModifyItemInventoryFields = ""

        ' make sure we have objects in our parameters
        If (itemInfoClass Is Nothing) Or (MainForm Is Nothing) Then
            Exit Function
        End If

        If (msgSetRequest Is Nothing) Then
            Exit Function
        End If
        
        '*** first set up modify request QB for this specific Item Inventory ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryMod As IItemInventoryMod
        Set itemInventoryMod = msgSetRequest.AppendItemInventoryModRq()

        ' specify the ListID of the item we want modified
        itemInventoryMod.ListID.SetValue (itemInfoClass.ListID)
        '<EditSequence>STRTYPE</EditSequence>
        itemInventoryMod.EditSequence.SetValue (MainForm.EditSequenceTextBox.Text)
        '<Name>STRTYPE</Name>                                <!-- opt, QBD max = 31 -->
        itemInventoryMod.name.SetValue (MainForm.NameTextBox.Text)
        '<IsActive>BOOLTYPE</IsActive>                       <!-- opt -->
        If (MainForm.IsActiveCheckBox.value = 0) Then
            itemInventoryMod.IsActive.SetValue (False)
        Else
            itemInventoryMod.IsActive.SetValue (True)
        End If
        '<ParentRef>                                         <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt -->
        '</ParentRef>
        listIndex = MainForm.ParentComboBox.listIndex
        Set objectInfo = Nothing
        If (listIndex >= 0) And (listIndex < itemCollection.count) Then
            Set objectInfo = itemCollection.Item(listIndex + 1)
        End If
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.ParentRef.ListID.SetValue (objectInfo.ListID)
        Else
            itemInventoryMod.ParentRef.ListID.SetEmpty
        End If
        '<SalesTaxCodeRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt, QBD max = 3 -->
        '</SalesTaxCodeRef>
        listIndex = MainForm.SalesTaxCodeComboBox.listIndex
        Set objectInfo = Nothing
        If (listIndex >= 0) And (listIndex < salesTaxCodeCollection.count) Then
            Set objectInfo = salesTaxCodeCollection.Item(listIndex + 1)
        End If
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.SalesTaxCodeRef.ListID.SetValue (objectInfo.ListID)
        Else
            itemInventoryMod.SalesTaxCodeRef.ListID.SetEmpty
        End If
        '<SalesDesc>STRTYPE</SalesDesc>                      <!-- opt, QBD max = 4095 -->
        itemInventoryMod.SalesDesc.SetValue (MainForm.SalesDescTextBox.Text)
        '<SalesPrice>PRICETYPE</SalesPrice>                  <!-- opt -->
        If (MainForm.SalesPriceTextBox.Text = "") Then
            itemInventoryMod.SalesPrice.SetEmpty
        Else
            itemInventoryMod.SalesPrice.SetValue (CDbl(MainForm.SalesPriceTextBox.Text))
        End If
        '<PurchaseDesc>STRTYPE</PurchaseDesc>                <!-- opt, QBD max = 4095 -->
        itemInventoryMod.PurchaseDesc.SetValue (MainForm.PurchaseDescTextBox.Text)
        '<PurchaseCost>PRICETYPE</PurchaseCost>              <!-- opt -->
        If (MainForm.PurchaseCostTextBox.Text = "") Then
            itemInventoryMod.PurchaseCost.SetEmpty
        Else
            itemInventoryMod.PurchaseCost.SetValue CDbl(MainForm.PurchaseCostTextBox.Text)
        End If
        '<COGSAccountRef>                                    <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>
        '</COGSAccountRef>
        listIndex = MainForm.COGSAccountComboBox.listIndex
        Set objectInfo = Nothing
        If (listIndex >= 0) And (listIndex < COGSCollection.count) Then
            Set objectInfo = COGSCollection.Item(listIndex + 1)
        End If
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.COGSAccountRef.ListID.SetValue (objectInfo.ListID)
        Else
            itemInventoryMod.COGSAccountRef.ListID.SetEmpty
        End If
        '<PrefVendorRef>                                     <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>
        '</PrefVendorRef>
        listIndex = MainForm.PrefVendorComboBox.listIndex
        Set objectInfo = Nothing
        If (listIndex >= 0) And (listIndex < vendorCollection.count) Then
            Set objectInfo = vendorCollection.Item(listIndex + 1)
        End If
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.PrefVendorRef.ListID.SetValue (objectInfo.ListID)
        Else
            itemInventoryMod.PrefVendorRef.ListID.SetEmpty
        End If
        '<AssetAccountRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>
        '</AssetAccountRef>
        listIndex = MainForm.AssetAccountComboBox.listIndex
        Set objectInfo = Nothing
        If (listIndex >= 0) And (listIndex < assetAccountCollection.count) Then
            Set objectInfo = assetAccountCollection.Item(listIndex + 1)
        End If
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.AssetAccountRef.ListID.SetValue (objectInfo.ListID)
        Else
            itemInventoryMod.AssetAccountRef.ListID.SetEmpty
        End If
        '<ReorderPoint>QUANTYPE</ReorderPoint>               <!-- opt -->
        If (MainForm.ReorderPointTextBox.Text = "") Then
            itemInventoryMod.ReorderPoint.SetEmpty
        Else
            itemInventoryMod.ReorderPoint.SetValue CDbl(MainForm.ReorderPointTextBox.Text)
        End If

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        'MsgBox msgSetRequest.ToXMLString(), vbOKOnly, "Request XML"
        'MsgBox msgSetResponse.ToXMLString(), vbOKOnly, "Response XML"

        '*** Now go thru the response from QB and add the Item Inventory fields ***
        '*** to the controls of the dialog box ***

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Function
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            Exit Function
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            Exit Function
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("ModifyItemInventoryFields unexpexcted Error - " & vbCrLf & "Status Code = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Function
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Function
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Function
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Function
        End If

        ' make sure we are processing the ItemInventoryModRs and
        ' the ItemInventoryRet responses in this response list
        Dim itemInventoryRet As IItemInventoryRet
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtItemInventoryModRs) Then
                ' save the response detail in the appropriate object type
                ' since we have first verified the type of the response object
            If (responseDetailType = ENObjectType.otItemInventoryRet) Then
                Set itemInventoryRet = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                Exit Function
            End If
        Else
            ' bail, we do not have the responses we were expecting
            Exit Function
        End If

        ModifyItemInventoryFields = DisplayItemInventoryFields(itemInventoryRet, MainForm)

        Exit Function
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in ModifyItemInventoryFields"
    End Function

    Public Function DisplayItemInventoryFields(ByRef itemInventoryRet As IItemInventoryRet, _
                        ByRef MainForm As MainForm) As String
        On Error GoTo Errs
        
        DisplayItemInventoryFields = ""

        ' make sure we have objects in our parameters
        If (itemInventoryRet Is Nothing) Or (MainForm Is Nothing) Then
            Exit Function
        End If

        ' clear the values in the form
        MainForm.ClearControls

        'Parse the query response and add the ItemInventory values to the field controls
                
        '<ListID>IDTYPE</ListID>
        '<EditSequence>STRTYPE</EditSequence>
        If (Not itemInventoryRet.EditSequence Is Nothing) Then
            MainForm.EditSequenceTextBox.Text = itemInventoryRet.EditSequence.GetValue()
        End If
        '<Name>STRTYPE</Name>                                <!-- opt, QBD max = 31 -->
        If (Not itemInventoryRet.name Is Nothing) Then
            MainForm.NameTextBox.Text = itemInventoryRet.name.GetValue()
        End If
        '<IsActive>BOOLTYPE</IsActive>                       <!-- opt -->
        If (Not itemInventoryRet.IsActive Is Nothing) Then
            If (itemInventoryRet.IsActive.GetValue()) Then
                MainForm.IsActiveCheckBox.value = 1
            Else
                MainForm.IsActiveCheckBox.value = 0
            End If
        End If
        '<ParentRef>                                         <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt -->
        '</ParentRef>
        If (Not itemInventoryRet.ParentRef Is Nothing) Then
            MainForm.ParentComboBox.listIndex = FindFullNameInComboBox(MainForm.ParentComboBox, itemInventoryRet.ParentRef.FullName.GetValue())
        End If
        '<SalesTaxCodeRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt, QBD max = 3 -->
        '</SalesTaxCodeRef>
        If (Not itemInventoryRet.SalesTaxCodeRef Is Nothing) Then
            MainForm.SalesTaxCodeComboBox.listIndex = FindFullNameInComboBox(MainForm.SalesTaxCodeComboBox, itemInventoryRet.SalesTaxCodeRef.FullName.GetValue())
        End If
        '<SalesDesc>STRTYPE</SalesDesc>                      <!-- opt, QBD max = 4095 -->
        If (Not itemInventoryRet.SalesDesc Is Nothing) Then
            MainForm.SalesDescTextBox.Text = itemInventoryRet.SalesDesc.GetValue()
        End If
        '<SalesPrice>PRICETYPE</SalesPrice>                  <!-- opt -->
        If (Not itemInventoryRet.SalesPrice Is Nothing) Then
            MainForm.SalesPriceTextBox.Text = itemInventoryRet.SalesPrice.GetValue()
        End If
        '<PurchaseDesc>STRTYPE</PurchaseDesc>                <!-- opt, QBD max = 4095 -->
        If (Not itemInventoryRet.PurchaseDesc Is Nothing) Then
            MainForm.PurchaseDescTextBox.Text = itemInventoryRet.PurchaseDesc.GetValue()
        End If
        '<PurchaseCost>PRICETYPE</PurchaseCost>              <!-- opt -->
        If (Not itemInventoryRet.PurchaseCost Is Nothing) Then
            MainForm.PurchaseCostTextBox.Text = itemInventoryRet.PurchaseCost.GetValue()
        End If
        '<COGSAccountRef>                                    <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>
        '</COGSAccountRef>
        If (Not itemInventoryRet.COGSAccountRef Is Nothing) Then
            MainForm.COGSAccountComboBox.listIndex = FindFullNameInComboBox(MainForm.COGSAccountComboBox, itemInventoryRet.COGSAccountRef.FullName.GetValue())
        End If
        '<PrefVendorRef>                                     <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>
        '</PrefVendorRef>
        If (Not itemInventoryRet.PrefVendorRef Is Nothing) Then
            MainForm.PrefVendorComboBox.listIndex = FindFullNameInComboBox(MainForm.PrefVendorComboBox, itemInventoryRet.PrefVendorRef.FullName.GetValue())
        End If
        '<AssetAccountRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>
        '</AssetAccountRef>
        If (Not itemInventoryRet.AssetAccountRef Is Nothing) Then
            MainForm.AssetAccountComboBox.listIndex = FindFullNameInComboBox(MainForm.AssetAccountComboBox, itemInventoryRet.AssetAccountRef.FullName.GetValue())
        End If
        '<ReorderPoint>QUANTYPE</ReorderPoint>               <!-- opt -->
        If (Not itemInventoryRet.ReorderPoint Is Nothing) Then
            MainForm.ReorderPointTextBox.Text = itemInventoryRet.ReorderPoint.GetValue()
        End If
        ' return the FullName
        If (Not itemInventoryRet.FullName Is Nothing) Then
            DisplayItemInventoryFields = itemInventoryRet.FullName.GetValue
        End If

        Exit Function
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in DisplayItemInventoryFields"
    End Function

    Public Sub FillSalesTaxCodeList(ByRef SalesTaxCodeComboBox As comboBox)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If
        
        '*** first query QB for the all of SalesTaxCode List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the SalesTaxCodeQuery request
        Dim salesTaxCodeQuery As ISalesTaxCodeQuery
        Set salesTaxCodeQuery = msgSetRequest.AppendSalesTaxCodeQueryRq()

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the SalesTaxCode names ***
        '*** to the combo box ***

        'Clear the combo box
        SalesTaxCodeComboBox.Clear
        If (salesTaxCodeCollection Is Nothing) Then
            Set salesTaxCodeCollection = New Collection
        End If
        Do While (salesTaxCodeCollection.count > 0)
            salesTaxCodeCollection.Remove (1)
        Loop

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("FillSalesTaxCodeList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the SalesTaxCodeQueryRs and
        ' the SalesTaxCodeRetList responses in this response list
        Dim salesTaxCodeRetList As ISalesTaxCodeRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtSalesTaxCodeQueryRs) And _
            (responseDetailType = ENObjectType.otSalesTaxCodeRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set salesTaxCodeRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Item to the SalesTaxCode combo box
        Dim count As Integer
        Dim index As Integer
        Dim salesTaxCodeRet As ISalesTaxCodeRet
        Dim itemInfo As FullNameListIDClass
        count = salesTaxCodeRetList.count
        For index = 0 To count - 1
            Set salesTaxCodeRet = salesTaxCodeRetList.GetAt(index)
            If (Not salesTaxCodeRet Is Nothing) Then
                If (Not salesTaxCodeRet.name Is Nothing) And _
                    (Not salesTaxCodeRet.ListID Is Nothing) Then
                    ' we are saving the FullName and ListID pairs for each SalesTaxCode
                    ' this is good practice since the FullName can change but the
                    ' ListID will not change for an SalesTaxCode.
                    Set itemInfo = New FullNameListIDClass
                    itemInfo.FullName = salesTaxCodeRet.name.GetValue()
                    itemInfo.ListID = salesTaxCodeRet.ListID.GetValue()
                    salesTaxCodeCollection.Add itemInfo, itemInfo.ListID
                    SalesTaxCodeComboBox.AddItem itemInfo.FullName
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillSalesTaxCodeListBox"
    End Sub

    Private Function FoundListIDInCollection(ByRef coll As Collection, _
                ByRef ListID As String) As Boolean
        On Error GoTo Errs

        FoundInComboBox = False

        Dim objectInfo As FullNameListIDClass

        ' go thru our combo box and find if any have this listID
        Dim i As Integer
        Dim numItems As Integer
        numItems = coll.count
        For i = 1 To numItems
            Set objectInfo = coll.Item(i)
            If (objectInfo.ListID = ListID) Then
                FoundListIDInCollection = True
                Exit For
            End If
        Next
        Exit Function
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FoundListIDInCollection"

    End Function
    Private Function FindFullNameInComboBox(ByRef comboBox As comboBox, _
                ByRef FullName As String) As Integer
        On Error GoTo Errs

        FindFullNameInComboBox = -1

        Dim name As String

        ' go thru our combo box and find the index which has this full name
        Dim i As Integer
        Dim numItems As Integer
        numItems = comboBox.ListCount
        For i = 0 To numItems - 1
            name = comboBox.List(i)
            If (name = FullName) Then
                FindFullNameInComboBox = i
                Exit Function
            End If
        Next

        MsgBox ("Error - FullName not found in the Combo Box = " & FullName & vbCrLf & "Probably need to reload the combo boxes.")
        Exit Function
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FindFullNameInComboBox"

    End Function

    Public Sub FillCOGSAccountList(ByRef COGSAccountComboBox As comboBox)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If
        
        '*** first query QB for the all of COGSAccount List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the AccountQuery request
        Dim accountQuery As IAccountQuery
        Set accountQuery = msgSetRequest.AppendAccountQueryRq()

        ' set we only want to return the COGS accounts
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add (ENAccountType.atCostOfGoodsSold)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the COGSAccount names ***
        '*** to the combo box ***

        'Clear the combo box
        COGSAccountComboBox.Clear
        If (COGSCollection Is Nothing) Then
            Set COGSCollection = New Collection
        End If
        Do While (COGSCollection.count > 0)
            COGSCollection.Remove (1)
        Loop

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("FillCOGSAccountList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the AccountQueryRs and
        ' the AccountRetList responses in this response list
        Dim accountRetList As IAccountRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtAccountQueryRs) And _
            (responseDetailType = ENObjectType.otAccountRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set accountRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Item to the COGSAccount box
        Dim count As Integer
        Dim index As Integer
        Dim accountRet As IAccountRet
        Dim itemInfo As FullNameListIDClass
        count = accountRetList.count
        For index = 0 To count - 1
            Set accountRet = accountRetList.GetAt(index)
            If (Not accountRet Is Nothing) Then
                If (Not accountRet.name Is Nothing) And _
                    (Not accountRet.AccountType Is Nothing) And _
                    (Not accountRet.ListID Is Nothing) Then
                    ' we are saving the FullName and ListID pairs for each Account
                    ' this is good practice since the FullName can change but the
                    ' ListID will not change for an Account.
                    If (accountRet.AccountType.GetValue() = ENAccountType.atCostOfGoodsSold) Then
                        Set itemInfo = New FullNameListIDClass
                        itemInfo.FullName = accountRet.name.GetValue()
                        itemInfo.ListID = accountRet.ListID.GetValue()
                        COGSCollection.Add itemInfo, itemInfo.ListID
                        COGSAccountComboBox.AddItem itemInfo.FullName
                    End If
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillCOGSAccountList"
    End Sub

    Public Sub FillAssetAccountList(ByRef AssetAccountComboBox As comboBox)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If
        
        '*** first query QB for the all of AssetAccount List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the AccountQuery request
        Dim accountQuery As IAccountQuery
        Set accountQuery = msgSetRequest.AppendAccountQueryRq()

        ' set we only want to return the AssetAccounts
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add (ENAccountType.atFixedAsset)
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add (ENAccountType.atOtherAsset)
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add (ENAccountType.atOtherCurrentAsset)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the AssetAccounts names ***
        '*** to the combo box ***

        'Clear the combo box and collection
        AssetAccountComboBox.Clear
        If (assetAccountCollection Is Nothing) Then
            Set assetAccountCollection = New Collection
        End If
        Do While (assetAccountCollection.count > 0)
            assetAccountCollection.Remove (1)
        Loop

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList Is Nothing) Then
            Exit Sub
        End If
        If (msgSetResponse.responseList.count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        Set responseList = msgSetResponse.responseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        Set response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox ("FillAssetAccountList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Then
            Exit Sub
        End If
        If (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Then
            Exit Sub
        End If
        If (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the AccountQueryRs and
        ' the AccountRetList responses in this response list
        Dim accountRetList As IAccountRetList
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtAccountQueryRs) And _
            (responseDetailType = ENObjectType.otAccountRetList) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            Set accountRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Item to the AssetAccounts box
        Dim count As Integer
        Dim index As Integer
        Dim accountRet As IAccountRet
        Dim itemInfo As FullNameListIDClass
        count = accountRetList.count
        For index = 0 To count - 1
            Set accountRet = accountRetList.GetAt(index)
            If (Not accountRet Is Nothing) Then
                If (Not accountRet.name Is Nothing) And _
                    (Not accountRet.AccountType Is Nothing) And _
                    (Not accountRet.ListID Is Nothing) Then
                    ' we are saving the FullName and ListID pairs for each Account
                    ' this is good practice since the FullName can change but the
                    ' ListID will not change for an Account.
                    If (accountRet.AccountType.GetValue() = ENAccountType.atFixedAsset) Or _
                        (accountRet.AccountType.GetValue() = ENAccountType.atOtherAsset) Or _
                        (accountRet.AccountType.GetValue() = ENAccountType.atOtherCurrentAsset) Then
                        Set itemInfo = New FullNameListIDClass
                        itemInfo.FullName = accountRet.name.GetValue()
                        itemInfo.ListID = accountRet.ListID.GetValue()
                        AssetAccountComboBox.AddItem itemInfo.FullName
                        assetAccountCollection.Add itemInfo, itemInfo.ListID
                    End If
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillAssetAccountList"
    End Sub

    Public Sub FillPrefVendorList(ByRef PrefVendorComboBox As comboBox)
        On Error GoTo Errs

        If (msgSetRequest Is Nothing) Then
            Exit Sub
        End If
        
        '*** first query QB for the all of Vendor List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the VendorQuery request
        Dim vendorQuery As IVendorQuery
        Set vendorQuery = msgSetRequest.AppendVendorQueryRq()

        ' we are going to set the number of entries returned to limit the size of the return structure
        vendorQuery.ORVendorListQuery.VendorListFilter.MaxReturned.SetValue (MAX_RETURNED)

        Dim bDone As Boolean
        bDone = False
        Dim firstFullName As String
        firstFullName = "!"

        'Clear the list box
        PrefVendorComboBox.Clear
        If (vendorCollection Is Nothing) Then
            Set vendorCollection = New Collection
        End If
        Do While (vendorCollection.count > 0)
            vendorCollection.Remove (1)
        Loop

        Do While (Not bDone)

            ' start looking for vendor next in the list
            ' we will have one overlap
            vendorQuery.ORVendorListQuery.VendorListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue (firstFullName)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            Set msgSetResponse = QBSessionManager.DoRequests(msgSetRequest)

            '*** Now go thru the response from QB and add the Vendor names ***
            '*** to the combo box ***

            ' check to make sure we have objects to access first
            ' and that there are responses in the list
            If (msgSetResponse Is Nothing) Then
                Exit Sub
            End If
            If (msgSetResponse.responseList Is Nothing) Then
                Exit Sub
            End If
            If (msgSetResponse.responseList.count <= 0) Then
                Exit Sub
            End If

            ' Start parsing the response list
            Dim responseList As IResponseList
            Set responseList = msgSetResponse.responseList

            ' go thru each response and process the response.
            ' this example will only have one response in the list
            ' so we will look at index=0
            Dim response As IResponse
            Set response = responseList.GetAt(0)
            If (Not response Is Nothing) Then
                If response.StatusCode <> "0" Then
                    'If the status is bad, report it to the user
                    MsgBox ("FillPrefVendorList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                    Exit Sub
                End If
            End If

            ' first make sure we have a response object to handle
            If (response Is Nothing) Then
                Exit Sub
            End If
            If (response.Type Is Nothing) Or _
                (response.Detail Is Nothing) Then
                Exit Sub
            End If
            If (response.Detail.Type Is Nothing) Then
                Exit Sub
            End If

            ' make sure we are processing the VendorQueryRs and
            ' the VendorRetList responses in this response list
            Dim vendorRetList As IVendorRetList
            Dim responseType As ENResponseType
            Dim responseDetailType As ENObjectType
            responseType = response.Type.GetValue()
            responseDetailType = response.Detail.Type.GetValue()
            If (responseType = ENResponseType.rtVendorQueryRs) And _
                (responseDetailType = ENObjectType.otVendorRetList) Then
                ' save the response detail in the appropriate object type
                ' since we have first verified the type of the response object
                Set vendorRetList = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                Exit Sub
            End If

            'Parse the query response and add the name to the Vendor
            Dim count As Integer
            Dim index As Integer
            Dim vendorRet As IVendorRet
            Dim itemInfo As FullNameListIDClass
            count = vendorRetList.count

            ' we are done with the vendorQueries if we have not received the MaxReturned
            If (count < MAX_RETURNED) Then
                bDone = True
            End If

            For index = 0 To count - 1
                Set vendorRet = vendorRetList.GetAt(index)
                If (Not vendorRet Is Nothing) Then
                    If (Not vendorRet.name Is Nothing) And _
                        (Not vendorRet.ListID Is Nothing) Then
                        ' we are saving the FullName and ListID pairs for each Vendor
                        ' this is good practice since the FullName can change but the
                        ' ListID will not change for a Vendor.
                        If (index <> 0) Or (Not FoundListIDInCollection(vendorCollection, vendorRet.ListID.GetValue())) Then
                            Set itemInfo = New FullNameListIDClass
                            itemInfo.FullName = vendorRet.name.GetValue()
                            itemInfo.ListID = vendorRet.ListID.GetValue()
                            vendorCollection.Add itemInfo, itemInfo.ListID
                            PrefVendorComboBox.AddItem itemInfo.FullName
                        End If
                        firstFullName = vendorRet.name.GetValue()
                    End If
                End If
            Next
        Loop

        Exit Sub
Errs:
        MsgBox "HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    vbOKOnly, _
                    "Error in FillPrefVendorList"
    End Sub

Function QBFCLatestVersion(SessionManager As QBSessionManager) As String
    Dim strXMLVersions() As String
    'Should be able to use this, but there appears to be a bug that may cause 2.0 to be returned
    'when it should not.
    'strXMLVersions = SessionManager.QBXMLVersionsForSession
    
    Dim msgset As IMsgSetRequest
    'Use oldest version to ensure that we work with any QuickBooks (US)
    Set msgset = SessionManager.CreateMsgSetRequest("US", 1, 0)
    msgset.AppendHostQueryRq
    Dim QueryResponse As IMsgSetResponse
    Set QueryResponse = SessionManager.DoRequests(msgset)
    Dim response As IResponse
    
    ' The response list contains only one response,
    ' which corresponds to our single HostQuery request
    Set response = QueryResponse.responseList.GetAt(0)
    Dim HostResponse As IHostRet
    Set HostResponse = response.Detail
    Dim supportedVersions As IBSTRList
    Set supportedVersions = HostResponse.SupportedQBXMLVersionList
    
    Dim i As Long
    Dim vers As Double
    Dim LastVers As Double
    LastVers = 0
    For i = 0 To supportedVersions.count - 1
        vers = Val(supportedVersions.GetAt(i))
        If (vers > LastVers) Then
            LastVers = vers
            QBFCLatestVersion = supportedVersions.GetAt(i)
        End If
    Next i
End Function
