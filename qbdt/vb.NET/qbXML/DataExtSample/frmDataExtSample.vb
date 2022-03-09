Option Strict Off
Option Explicit On
Friend Class frmDataExtSample
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------------
	' Form: frmDataExtSample
	'
	' Description: This the main form and entry point for this sample
	'              program.  It displays the currently defined data
	'              extension definitions and custom fields for the open
	'              company file.  From here the user may choose to
	'              activate forms to add data extension definitions to the
	'              company file, define values for data extension
	'              for customers or modify data extension values for
	'              customers
	'
	'              The form calls OpenSessionBeginSession to make sure
	'              a company file is open.
	'
	' Copyright © 2021-2022 Intuit Inc. All rights reserved.
	' Use is subject to the terms specified at:
	'      http://developer.intuit.com/legal/devsite_tos.html
	'
	'----------------------------------------------------------
	Private Sub cmdAddDataExtension_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAddDataExtension.Click
        If CustomersHaveDataExts Then
            frmAddDataExtension.ShowDialog()
        Else
            MsgBox("You must add a data extension for Customers before adding a data extension value to a specified customer")
		End If
	End Sub

    Private Sub cmdDefineDataExt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDefineDataExt.Click
        frmAddDataExtDef.ShowDialog()
    End Sub

    Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		EndSessionCloseConnection()
		End
	End Sub
	
	Private Sub cmdModDataExt_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdModDataExt.Click
		If CustomersHaveDataExts Then
            frmModDataExtension.ShowDialog()
        Else
			MsgBox("You must add a data extension for Customers before modifying a data extension value for a specific customer")
		End If
	End Sub
	
	Private Sub cmdShowCustomRequest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowCustomRequest.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strCustomRequest))
        frmModDataExtension.ShowDialog()
    End Sub
	
	Private Sub cmdShowCustomResponse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowCustomResponse.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strCustomResponse))
        frmqbXMLDisplay.ShowDialog()
    End Sub
	
	Private Sub cmdShowQueryRequest_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowQueryRequest.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strQueryRequest))
        frmqbXMLDisplay.ShowDialog()
    End Sub
	
	Private Sub cmdShowQueryResponse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdShowQueryResponse.Click
		frmqbXMLDisplay.txtqbXML.Text = PrettyXMLString((modDataExtSample.strQueryResponse))
        frmqbXMLDisplay.ShowDialog()

    End Sub
	
	Private Sub frmDataExtSample_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		OpenConnectionBeginSession()
		GetDataExts(txtDataExts)
		GetCustomFields(txtCustomFields)
	End Sub
End Class