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
' Updated itemInventoryQuery.ORListQuery to new query 09/2012
'----------------------------------------------------------
Imports Interop.QBFC13


Module ItemInventoryModule
    Dim booSessionBegun As Boolean

    Dim MAX_RETURNED As Integer = 20

    'Module objects
    Dim qbSessionManager As qbSessionManager
    Dim msgSetRequest As IMsgSetRequest

    Public Sub OpenConnectionBeginSession()

        booSessionBegun = False

        On Error GoTo Errs

        qbSessionManager = New QBSessionManager()

        qbSessionManager.OpenConnection("", "IDN Item Inventory Modify Sample - QBFC")

        qbSessionManager.BeginSession("", ENOpenMode.omDontCare)
        booSessionBegun = True

        'Check to make sure the QuickBooks we're working with supports version 2 of the SDK
        Dim strXMLVersions() As String
        strXMLVersions = qbSessionManager.QBXMLVersionsForSession

        Dim booSupports2dot0 As Boolean
        booSupports2dot0 = False
        Dim i As Long
        For i = LBound(strXMLVersions) To UBound(strXMLVersions)
            If (strXMLVersions(i) = "2.0") Then
                booSupports2dot0 = True
                msgSetRequest = qbSessionManager.CreateMsgSetRequest("US", 2, 0)
                Exit For
            End If
        Next

        If Not booSupports2dot0 Then
            MsgBox("This program only runs against QuickBooks installations which support the 2.0 qbXML spec.  Your version of QuickBooks does not support qbXML 2.0")
            End
        End If

        Exit Sub

Errs:
        If Err.Number = &H80040416 Then
            MsgBox("You must have QuickBooks running with the company" & vbCrLf & "file open to use this program.")
            qbSessionManager.CloseConnection()
            End
        ElseIf Err.Number = &H80040422 Then
            MsgBox("This QuickBooks company file is open in single user mode and" & vbCrLf & "another application is already accessing it.  Please exit the" & vbCrLf & "other application and run this application again.")
            qbSessionManager.CloseConnection()
            End
        Else
            MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, MsgBoxStyle.Critical, "Error in OpenConnectionBeginSession")

            If booSessionBegun Then
                qbSessionManager.EndSession()
            End If

            qbSessionManager.CloseConnection()
            End
        End If
    End Sub

    Public Sub EndSessionCloseConnection()
        On Error GoTo Errs
        If booSessionBegun Then
            qbSessionManager.EndSession()
            qbSessionManager.CloseConnection()
        End If
        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in EndSessionCloseConnection")
    End Sub

    Public Sub FillItemInventoryListBox(ByRef itemInventoryList As System.Windows.Forms.ListBox, _
                        ByRef parentComboBox As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        '*** first query QB for the all of Item Inventory List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryQuery As IItemInventoryQuery
        itemInventoryQuery = msgSetRequest.AppendItemInventoryQueryRq()

        ' lets get all the Item Inventory, active and inactive.
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ActiveStatus.SetValue(ENActiveStatus.asAll)

        ' we are going to set the number of entries returned to limit the size of the return structure
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.MaxReturned.SetValue(MAX_RETURNED)

        Dim bDone As Boolean = False
        Dim firstFullName As String = "!"

        'Clear the list box
        itemInventoryList.Items.Clear()
        parentComboBox.Items.Clear()

        Do While (Not bDone)

            ' start looking for item next in the list
            ' we will have one overlap
            itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListWithClassFilter.ORNameFilter.NameRangeFilter.FromName.SetValue(firstFullName)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

            '*** Now go thru the response from QB and add the Item Inventory names ***
            '*** to the list box ***

            ' check to make sure we have objects to access first
            ' and that there are responses in the list
            If (msgSetResponse Is Nothing) Or _
                (msgSetResponse.ResponseList Is Nothing) Or _
                (msgSetResponse.ResponseList.Count <= 0) Then
                Exit Sub
            End If

            ' Start parsing the response list
            Dim responseList As IResponseList
            responseList = msgSetResponse.ResponseList

            ' go thru each response and process the response.
            ' this example will only have one response in the list
            ' so we will look at index=0
            Dim response As IResponse
            response = responseList.GetAt(0)
            If (Not response Is Nothing) Then
                If response.StatusCode <> "0" Then
                    'If the status is bad, report it to the user
                    MsgBox("FillItemInventoryListBox unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                    Exit Sub
                End If
            End If

            ' first make sure we have a response object to handle
            If (response Is Nothing) Or _
                (response.Type Is Nothing) Or _
                (response.Detail Is Nothing) Or _
                (response.Detail.Type Is Nothing) Then
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
                itemInventoryRetList = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                Exit Sub
            End If

            'Parse the query response and add the Item to the Item Inventory list box
            Dim count As Short
            Dim index As Short
            Dim itemInventoryRet As IItemInventoryRet
            Dim itemInfo As FullNameListIDClass
            count = itemInventoryRetList.Count

            ' we are done with the itemInventoryQueries if we have not received the MaxReturned
            If (count < MAX_RETURNED) Then
                bDone = True
            End If

            For index = 0 To count - 1
                itemInventoryRet = itemInventoryRetList.GetAt(index)
                If (Not itemInventoryRet Is Nothing) And _
                    (Not itemInventoryRet.FullName Is Nothing) And _
                    (Not itemInventoryRet.ListID Is Nothing) Then
                    ' we are saving the FullName and ListID pairs for each Item Inventory
                    ' this is good practice since the FullName can change but the 
                    ' ListID will not change for an Item Inventory.
                    itemInfo = New FullNameListIDClass()
                    itemInfo.FullName = itemInventoryRet.FullName.GetValue()
                    itemInfo.ListID = itemInventoryRet.ListID.GetValue()
                    If (index <> 0) Or (Not FoundInComboBox(parentComboBox, itemInventoryRet.ListID.GetValue())) Then
                        ' if the item is not in the combo box, then it is also not in the list box
                        itemInventoryList.Items.Add(itemInfo)
                        parentComboBox.Items.Add(itemInfo)
                    End If
                    firstFullName = itemInventoryRet.FullName.GetValue()
                End If
            Next
        Loop


        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillItemInventoryListBox")
    End Sub

    Public Sub QueryItemInventoryFields(ByRef itemInfoClass As FullNameListIDClass, _
                        ByRef mainForm As MainForm)
        On Error GoTo Errs

        '*** first query QB for this specific Item Inventory ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryQuery As IItemInventoryQuery
        itemInventoryQuery = msgSetRequest.AppendItemInventoryQueryRq()

        ' specify the ListID of the item we want returned
        itemInventoryQuery.ORListQueryWithOwnerIDAndClass.ListIDList.Add(itemInfoClass.ListID)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the Item Inventory fields ***
        '*** to the controls of the dialog box ***

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("QueryItemInventoryFields unexpexcted Error - " & vbCrLf & "Status Code = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
            itemInventoryRetList = response.Detail
            If (itemInventoryRetList.Count >= 0) Then
                itemInventoryRet = itemInventoryRetList.GetAt(0)
            End If
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        DisplayItemInventoryFields(itemInventoryRet, mainForm)

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in QueryItemInventoryListBox")
    End Sub

    Public Sub ModifyItemInventoryFields(ByRef itemInfoClass As FullNameListIDClass, _
                        ByRef mainForm As MainForm)
        On Error GoTo Errs

        Dim objectInfo As FullNameListIDClass

        ' make sure we have objects in our parameters
        If (itemInfoClass Is Nothing) Or (mainForm Is Nothing) Then
            Exit Sub
        End If

        '*** first set up modify request QB for this specific Item Inventory ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the ItemInventoryQuery request
        Dim itemInventoryMod As IItemInventoryMod
        itemInventoryMod = msgSetRequest.AppendItemInventoryModRq()

        ' specify the ListID of the item we want modified
        itemInventoryMod.ListID.SetValue(itemInfoClass.ListID)
        '<EditSequence>STRTYPE</EditSequence>                
        itemInventoryMod.EditSequence.SetValue(mainForm.EditSequenceTextBox.Text)
        '<Name>STRTYPE</Name>                                <!-- opt, QBD max = 31 -->
        itemInventoryMod.Name.SetValue(mainForm.NameTextBox.Text)
        '<IsActive>BOOLTYPE</IsActive>                       <!-- opt -->
        itemInventoryMod.IsActive.SetValue(mainForm.IsActiveCheckBox.Checked)
        '<ParentRef>                                         <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt -->
        '</ParentRef>
        objectInfo = mainForm.ParentComboBox.SelectedItem
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.ParentRef.ListID.SetValue(objectInfo.ListID)
        Else
            itemInventoryMod.ParentRef.ListID.SetEmpty()
        End If
        '<SalesTaxCodeRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt, QBD max = 3 -->
        '</SalesTaxCodeRef>
        objectInfo = mainForm.SalesTaxCodeComboBox.SelectedItem
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.SalesTaxCodeRef.ListID.SetValue(objectInfo.ListID)
        Else
            itemInventoryMod.SalesTaxCodeRef.ListID.SetEmpty()
        End If
        '<SalesDesc>STRTYPE</SalesDesc>                      <!-- opt, QBD max = 4095 -->
        itemInventoryMod.SalesDesc.SetValue(mainForm.SalesDescTextBox.Text)
        '<SalesPrice>PRICETYPE</SalesPrice>                  <!-- opt -->
        If (mainForm.SalesPriceTextBox.Text.Length = 0) Then
            itemInventoryMod.SalesPrice.SetEmpty()
        Else
            itemInventoryMod.SalesPrice.SetValue(CDbl(mainForm.SalesPriceTextBox.Text))
        End If
        '<PurchaseDesc>STRTYPE</PurchaseDesc>                <!-- opt, QBD max = 4095 -->
        itemInventoryMod.PurchaseDesc.SetValue(mainForm.PurchaseDescTextBox.Text)
        '<PurchaseCost>PRICETYPE</PurchaseCost>              <!-- opt -->
        If (mainForm.PurchaseCostTextBox.Text.Length = 0) Then
            itemInventoryMod.PurchaseCost.SetEmpty()
        Else
            itemInventoryMod.PurchaseCost.SetValue(CDbl(mainForm.PurchaseCostTextBox.Text))
        End If
        '<COGSAccountRef>                                    <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      
        '</COGSAccountRef>
        objectInfo = mainForm.COGSAccountComboBox.SelectedItem
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.COGSAccountRef.ListID.SetValue(objectInfo.ListID)
        Else
            itemInventoryMod.COGSAccountRef.ListID.SetEmpty()
        End If
        '<PrefVendorRef>                                     <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      
        '</PrefVendorRef>
        objectInfo = mainForm.PrefVendorComboBox.SelectedItem
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.PrefVendorRef.ListID.SetValue(objectInfo.ListID)
        Else
            itemInventoryMod.PrefVendorRef.ListID.SetEmpty()
        End If
        '<AssetAccountRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      
        '</AssetAccountRef>
        objectInfo = mainForm.AssetAccountComboBox.SelectedItem
        If (Not objectInfo Is Nothing) Then
            itemInventoryMod.AssetAccountRef.ListID.SetValue(objectInfo.ListID)
        Else
            itemInventoryMod.AssetAccountRef.ListID.SetEmpty()
        End If
        '<ReorderPoint>QUANTYPE</ReorderPoint>               <!-- opt -->
        If (mainForm.ReorderPointTextBox.Text.Length = 0) Then
            itemInventoryMod.ReorderPoint.SetEmpty()
        Else
            itemInventoryMod.ReorderPoint.SetValue(CDbl(mainForm.ReorderPointTextBox.Text))
        End If

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        'MsgBox(msgSetRequest.ToXMLString(), MsgBoxStyle.Information, "Request XML")
        'MsgBox(msgSetResponse.ToXMLString(), MsgBoxStyle.Information, "Response XML")

        '*** Now go thru the response from QB and add the Item Inventory fields ***
        '*** to the controls of the dialog box ***

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("ModifyItemInventoryFields unexpexcted Error - " & vbCrLf & "Status Code = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
            Exit Sub
        End If

        ' make sure we are processing the ItemInventoryModRs and 
        ' the ItemInventoryRet responses in this response list
        Dim itemInventoryRet As IItemInventoryRet
        Dim responseType As ENResponseType
        Dim responseDetailType As ENObjectType
        responseType = response.Type.GetValue()
        responseDetailType = response.Detail.Type.GetValue()
        If (responseType = ENResponseType.rtItemInventoryModRs) And _
            (responseDetailType = ENObjectType.otItemInventoryRet) Then
            ' save the response detail in the appropriate object type
            ' since we have first verified the type of the response object
            itemInventoryRet = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        DisplayItemInventoryFields(itemInventoryRet, mainForm)

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in ModifyItemInventoryFields")
    End Sub

    Public Sub DisplayItemInventoryFields(ByRef itemInventoryRet As IItemInventoryRet, _
                        ByRef mainForm As MainForm)
        On Error GoTo Errs

        ' make sure we have objects in our parameters
        If (itemInventoryRet Is Nothing) Or (mainForm Is Nothing) Then
            Exit Sub
        End If

        ' clear the values in the form
        mainForm.ClearControls()

        'Parse the query response and add the ItemInventory values to the field controls
        '<ListID>IDTYPE</ListID>
        '<EditSequence>STRTYPE</EditSequence>                
        If (Not itemInventoryRet.EditSequence Is Nothing) Then
            mainForm.EditSequenceTextBox.Text = itemInventoryRet.EditSequence.GetValue()
        End If
        '<Name>STRTYPE</Name>                                <!-- opt, QBD max = 31 -->
        If (Not itemInventoryRet.Name Is Nothing) Then
            mainForm.NameTextBox.Text = itemInventoryRet.Name.GetValue()
        End If
        '<IsActive>BOOLTYPE</IsActive>                       <!-- opt -->
        If (Not itemInventoryRet.IsActive Is Nothing) Then
            mainForm.IsActiveCheckBox.Checked = itemInventoryRet.IsActive.GetValue()
        End If
        '<ParentRef>                                         <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt -->
        '</ParentRef>
        If (Not itemInventoryRet.ParentRef Is Nothing) Then
            mainForm.ParentComboBox.SelectedIndex = FindFullNameInComboBox(mainForm.ParentComboBox, itemInventoryRet.ParentRef.FullName.GetValue())
        End If
        '<SalesTaxCodeRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      <!-- opt, QBD max = 3 -->
        '</SalesTaxCodeRef>
        If (Not itemInventoryRet.SalesTaxCodeRef Is Nothing) Then
            mainForm.SalesTaxCodeComboBox.SelectedIndex = FindFullNameInComboBox(mainForm.SalesTaxCodeComboBox, itemInventoryRet.SalesTaxCodeRef.FullName.GetValue())
        End If
        '<SalesDesc>STRTYPE</SalesDesc>                      <!-- opt, QBD max = 4095 -->
        If (Not itemInventoryRet.SalesDesc Is Nothing) Then
            mainForm.SalesDescTextBox.Text = itemInventoryRet.SalesDesc.GetValue()
        End If
        '<SalesPrice>PRICETYPE</SalesPrice>                  <!-- opt -->
        If (Not itemInventoryRet.SalesPrice Is Nothing) Then
            mainForm.SalesPriceTextBox.Text = itemInventoryRet.SalesPrice.GetValue()
        End If
        '<PurchaseDesc>STRTYPE</PurchaseDesc>                <!-- opt, QBD max = 4095 -->
        If (Not itemInventoryRet.PurchaseDesc Is Nothing) Then
            mainForm.PurchaseDescTextBox.Text = itemInventoryRet.PurchaseDesc.GetValue()
        End If
        '<PurchaseCost>PRICETYPE</PurchaseCost>              <!-- opt -->
        If (Not itemInventoryRet.PurchaseCost Is Nothing) Then
            mainForm.PurchaseCostTextBox.Text = itemInventoryRet.PurchaseCost.GetValue()
        End If
        '<COGSAccountRef>                                    <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      
        '</COGSAccountRef>
        If (Not itemInventoryRet.COGSAccountRef Is Nothing) Then
            mainForm.COGSAccountComboBox.SelectedIndex = FindFullNameInComboBox(mainForm.COGSAccountComboBox, itemInventoryRet.COGSAccountRef.FullName.GetValue())
        End If
        '<PrefVendorRef>                                     <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      
        '</PrefVendorRef>
        If (Not itemInventoryRet.PrefVendorRef Is Nothing) Then
            mainForm.PrefVendorComboBox.SelectedIndex = FindFullNameInComboBox(mainForm.PrefVendorComboBox, itemInventoryRet.PrefVendorRef.FullName.GetValue())
        End If
        '<AssetAccountRef>                                   <!-- opt -->
        '  <ListID>IDTYPE</ListID>                           <!-- opt -->
        '  <FullName>STRTYPE</FullName>                      
        '</AssetAccountRef>
        If (Not itemInventoryRet.AssetAccountRef Is Nothing) Then
            mainForm.AssetAccountComboBox.SelectedIndex = FindFullNameInComboBox(mainForm.AssetAccountComboBox, itemInventoryRet.AssetAccountRef.FullName.GetValue())
        End If
        '<ReorderPoint>QUANTYPE</ReorderPoint>               <!-- opt -->
        If (Not itemInventoryRet.ReorderPoint Is Nothing) Then
            mainForm.ReorderPointTextBox.Text = itemInventoryRet.ReorderPoint.GetValue()
        End If

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in DisplayItemInventoryFields")
    End Sub

    Public Sub FillSalesTaxCodeList(ByRef salesTaxCodeComboBox As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        '*** first query QB for the all of SalesTaxCode List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the SalesTaxCodeQuery request
        Dim salesTaxCodeQuery As ISalesTaxCodeQuery
        salesTaxCodeQuery = msgSetRequest.AppendSalesTaxCodeQueryRq()

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the SalesTaxCode names ***
        '*** to the combo box ***

        'Clear the combo box
        salesTaxCodeComboBox.Items.Clear()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("FillSalesTaxCodeList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
            salesTaxCodeRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Item to the SalesTaxCode combo box
        Dim count As Short
        Dim index As Short
        Dim salesTaxCodeRet As ISalesTaxCodeRet
        Dim itemInfo As FullNameListIDClass
        count = salesTaxCodeRetList.Count
        For index = 0 To count - 1
            salesTaxCodeRet = salesTaxCodeRetList.GetAt(index)
            If (Not salesTaxCodeRet Is Nothing) And _
                (Not salesTaxCodeRet.Name Is Nothing) And _
                (Not salesTaxCodeRet.ListID Is Nothing) Then
                ' we are saving the FullName and ListID pairs for each SalesTaxCode
                ' this is good practice since the FullName can change but the 
                ' ListID will not change for an SalesTaxCode.
                itemInfo = New FullNameListIDClass()
                itemInfo.FullName = salesTaxCodeRet.Name.GetValue()
                itemInfo.ListID = salesTaxCodeRet.ListID.GetValue()
                salesTaxCodeComboBox.Items.Add(itemInfo)
            End If
        Next

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillSalesTaxCodeListBox")
    End Sub

    Private Function FoundInComboBox(ByRef comboBox As System.Windows.Forms.ComboBox, _
                ByRef listID As String) As Boolean
        On Error GoTo Errs

        FoundInComboBox = False

        Dim objectInfo As FullNameListIDClass

        ' go thru our combo box and find if any have this listID
        Dim i As Short
        Dim numCustomers As Short
        numCustomers = comboBox.Items.Count
        For i = 0 To numCustomers - 1
            objectInfo = comboBox.Items.Item(i)
            If (objectInfo.ListID = listID) Then
                FoundInComboBox = True
                Exit For
            End If
        Next
        Exit Function
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FoundInComboBox")

    End Function
    Private Function FindFullNameInComboBox(ByRef comboBox As System.Windows.Forms.ComboBox, _
                ByRef fullName As String) As Integer
        On Error GoTo Errs

        FindFullNameInComboBox = -1

        Dim objectInfo As FullNameListIDClass

        ' go thru our combo box and find the index which has this full name
        Dim i As Integer
        Dim numCustomers As Integer
        numCustomers = comboBox.Items.Count
        For i = 0 To numCustomers - 1
            objectInfo = comboBox.Items.Item(i)
            If (objectInfo.FullName = fullName) Then
                FindFullNameInComboBox = i
                Exit Function
            End If
        Next

        MsgBox("Error - FullName not found in the Combo Box = " & fullName & vbCrLf & "Probably need to reload the combo boxes.")
        Exit Function
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FindFullNameInComboBox")

    End Function

    Public Sub FillCOGSAccountList(ByRef COGSAccountComboBox As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        '*** first query QB for the all of COGSAccount List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the AccountQuery request
        Dim accountQuery As IAccountQuery
        accountQuery = msgSetRequest.AppendAccountQueryRq()

        ' set we only want to return the COGS accounts
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(ENAccountType.atCostOfGoodsSold)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the COGSAccount names ***
        '*** to the combo box ***

        'Clear the combo box
        COGSAccountComboBox.Items.Clear()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("FillCOGSAccountList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
            accountRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Item to the COGSAccount box
        Dim count As Short
        Dim index As Short
        Dim accountRet As IAccountRet
        Dim itemInfo As FullNameListIDClass
        count = accountRetList.Count
        For index = 0 To count - 1
            accountRet = accountRetList.GetAt(index)
            If (Not accountRet Is Nothing) And _
                (Not accountRet.Name Is Nothing) And _
                (Not accountRet.AccountType Is Nothing) And _
                (Not accountRet.ListID Is Nothing) Then
                ' we are saving the FullName and ListID pairs for each Account
                ' this is good practice since the FullName can change but the 
                ' ListID will not change for an Account.
                If (accountRet.AccountType.GetValue() = ENAccountType.atCostOfGoodsSold) Then
                    itemInfo = New FullNameListIDClass()
                    itemInfo.FullName = accountRet.Name.GetValue()
                    itemInfo.ListID = accountRet.ListID.GetValue()
                    COGSAccountComboBox.Items.Add(itemInfo)
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillCOGSAccountList")
    End Sub

    Public Sub FillAssetAccountList(ByRef AssetAccountComboBox As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        '*** first query QB for the all of AssetAccount List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the AccountQuery request
        Dim accountQuery As IAccountQuery
        accountQuery = msgSetRequest.AppendAccountQueryRq()

        ' set we only want to return the AssetAccounts
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(ENAccountType.atFixedAsset)
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(ENAccountType.atOtherAsset)
        accountQuery.ORAccountListQuery.AccountListFilter.AccountTypeList.Add(ENAccountType.atOtherCurrentAsset)

        ' send the request to QB
        Dim msgSetResponse As IMsgSetResponse
        msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

        '*** Now go thru the response from QB and add the AssetAccounts names ***
        '*** to the combo box ***

        'Clear the combo box
        AssetAccountComboBox.Items.Clear()

        ' check to make sure we have objects to access first
        ' and that there are responses in the list
        If (msgSetResponse Is Nothing) Or _
            (msgSetResponse.ResponseList Is Nothing) Or _
            (msgSetResponse.ResponseList.Count <= 0) Then
            Exit Sub
        End If

        ' Start parsing the response list
        Dim responseList As IResponseList
        responseList = msgSetResponse.ResponseList

        ' go thru each response and process the response.
        ' this example will only have one response in the list
        ' so we will look at index=0
        Dim response As IResponse
        response = responseList.GetAt(0)
        If (Not response Is Nothing) Then
            If response.StatusCode <> "0" Then
                'If the status is bad, report it to the user
                MsgBox("FillAssetAccountList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                Exit Sub
            End If
        End If

        ' first make sure we have a response object to handle
        If (response Is Nothing) Or _
            (response.Type Is Nothing) Or _
            (response.Detail Is Nothing) Or _
            (response.Detail.Type Is Nothing) Then
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
            accountRetList = response.Detail
        Else
            ' bail, we do not have the responses we were expecting
            Exit Sub
        End If

        'Parse the query response and add the Item to the AssetAccounts box
        Dim count As Short
        Dim index As Short
        Dim accountRet As IAccountRet
        Dim itemInfo As FullNameListIDClass
        count = accountRetList.Count
        For index = 0 To count - 1
            accountRet = accountRetList.GetAt(index)
            If (Not accountRet Is Nothing) And _
                (Not accountRet.Name Is Nothing) And _
                (Not accountRet.AccountType Is Nothing) And _
                (Not accountRet.ListID Is Nothing) Then
                ' we are saving the FullName and ListID pairs for each Account
                ' this is good practice since the FullName can change but the 
                ' ListID will not change for an Account.
                If (accountRet.AccountType.GetValue() = ENAccountType.atFixedAsset) Or _
                    (accountRet.AccountType.GetValue() = ENAccountType.atOtherAsset) Or _
                    (accountRet.AccountType.GetValue() = ENAccountType.atOtherCurrentAsset) Then
                    itemInfo = New FullNameListIDClass()
                    itemInfo.FullName = accountRet.Name.GetValue()
                    itemInfo.ListID = accountRet.ListID.GetValue()
                    AssetAccountComboBox.Items.Add(itemInfo)
                End If
            End If
        Next

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillAssetAccountList")
    End Sub

    Public Sub FillPrefVendorList(ByRef prefVendorComboBox As System.Windows.Forms.ComboBox)
        On Error GoTo Errs

        '*** first query QB for the all of Vendor List ***

        ' make sure we do not have any old requests still defined
        msgSetRequest.ClearRequests()

        ' set the OnError attribute to continueOnError
        msgSetRequest.Attributes.OnError = ENRqOnError.roeContinue

        ' Add the VendorQuery request
        Dim vendorQuery As IVendorQuery
        vendorQuery = msgSetRequest.AppendVendorQueryRq()

        ' we are going to set the number of entries returned to limit the size of the return structure
        vendorQuery.ORVendorListQuery.VendorListFilter.MaxReturned.SetValue(MAX_RETURNED)

        Dim bDone As Boolean = False
        Dim firstFullName As String = "!"

        'Clear the list box
        prefVendorComboBox.Items.Clear()

        Do While (Not bDone)

            ' start looking for vendor next in the list
            ' we will have one overlap
            vendorQuery.ORVendorListQuery.VendorListFilter.ORNameFilter.NameRangeFilter.FromName.SetValue(firstFullName)

            ' send the request to QB
            Dim msgSetResponse As IMsgSetResponse
            msgSetResponse = qbSessionManager.DoRequests(msgSetRequest)

            '*** Now go thru the response from QB and add the Vendor names ***
            '*** to the combo box ***

            ' check to make sure we have objects to access first
            ' and that there are responses in the list
            If (msgSetResponse Is Nothing) Or _
                (msgSetResponse.ResponseList Is Nothing) Or _
                (msgSetResponse.ResponseList.Count <= 0) Then
                Exit Sub
            End If

            ' Start parsing the response list
            Dim responseList As IResponseList
            responseList = msgSetResponse.ResponseList

            ' go thru each response and process the response.
            ' this example will only have one response in the list
            ' so we will look at index=0
            Dim response As IResponse
            response = responseList.GetAt(0)
            If (Not response Is Nothing) Then
                If response.StatusCode <> "0" Then
                    'If the status is bad, report it to the user
                    MsgBox("FillPrefVendorList unexpexcted Error - " & vbCrLf & "StatusCode = " & response.StatusCode & vbCrLf & vbCrLf & response.StatusMessage)
                    Exit Sub
                End If
            End If

            ' first make sure we have a response object to handle
            If (response Is Nothing) Or _
                (response.Type Is Nothing) Or _
                (response.Detail Is Nothing) Or _
                (response.Detail.Type Is Nothing) Then
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
                vendorRetList = response.Detail
            Else
                ' bail, we do not have the responses we were expecting
                Exit Sub
            End If

            'Parse the query response and add the name to the Vendor
            Dim count As Short
            Dim index As Short
            Dim vendorRet As IVendorRet
            Dim itemInfo As FullNameListIDClass
            count = vendorRetList.Count

            ' we are done with the vendorQueries if we have not received the MaxReturned
            If (count < MAX_RETURNED) Then
                bDone = True
            End If

            For index = 0 To count - 1
                vendorRet = vendorRetList.GetAt(index)
                If (Not vendorRet Is Nothing) And _
                    (Not vendorRet.Name Is Nothing) And _
                    (Not vendorRet.ListID Is Nothing) Then
                    ' we are saving the FullName and ListID pairs for each Vendor
                    ' this is good practice since the FullName can change but the 
                    ' ListID will not change for a Vendor.
                    itemInfo = New FullNameListIDClass()
                    itemInfo.FullName = vendorRet.Name.GetValue()
                    itemInfo.ListID = vendorRet.ListID.GetValue()
                    If (index <> 0) Or (Not FoundInComboBox(prefVendorComboBox, vendorRet.ListID.GetValue())) Then
                        prefVendorComboBox.Items.Add(itemInfo)
                    End If
                    firstFullName = vendorRet.Name.GetValue()
                End If
            Next
        Loop

        Exit Sub
Errs:
        MsgBox("HRESULT = " & Err.Number & " (" & Hex(Err.Number) & ") " & vbCrLf & vbCrLf & Err.Description, _
                    MsgBoxStyle.Critical, _
                    "Error in FillPrefVendorList")
    End Sub

End Module
