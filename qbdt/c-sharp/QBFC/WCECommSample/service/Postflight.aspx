<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Postflight.aspx.cs" Inherits="Postflight" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Post-flight</title>
    <script src="javascript/prototype.js" language="javascript" type="text/javascript"></script>
</head>
<body>
<script language="javascript" type="text/javascript">
function reportError (resp) {
    alert("Error " + resp.status + ': ' + resp.statusText);
}
function processData(request) {
   if(request.status == 200){
        //Success status=200
        //alert(request.responseText);
   }
}
function finalResult(request) {
	alert("Final result: " + request.status + ': ' + request.responseText);
}
var strSessionID;
var strListID;
var strTxnID;

function ListDisplay(strListID, strSessionID) {
    var command="qbwc://doquery/ListDisplayMod?ListID="+strListID+"&type=Customer&session=" + strSessionID+"&async=true";
    //alert(command);
    var req = new Ajax.Request(command, {method:'get', onSuccess:processData, onFailure:reportError});
}

function TxnDisplay(strTxnID, strSessionID) {
    var command="qbwc://doquery/TxnDisplayMod?TxnID="+strTxnID+"&type=SalesReceipt&session=" + strSessionID+"&async=true";
    //alert(command);
    var req = new Ajax.Request(command, {method:'get', onSuccess:processData, onFailure:reportError});
}

</script>

    <form id="form1" runat="server">
    <div>
        <strong><span style="font-size: 16pt">POST-FLIGHT INFORMATION<br />
            <br />
        </span></strong>The following customer and sales receipt data has been transferred to
        QuickBooks.<br />
        <br />
        <asp:Label ID="Label1" runat="server" Text="Customers"></asp:Label><br />
        <asp:Table ID="Table_Customers" runat="server" BorderStyle="Ridge" GridLines="Both"
            Width="797px">
        </asp:Table>
        &nbsp;&nbsp;&nbsp;<br />
        <asp:Label ID="Label2" runat="server" Text="Sales Receipts"></asp:Label><br />
        <asp:Table ID="Table_SalesReceipts" runat="server" BorderStyle="Ridge" GridLines="Both"
            Width="795px">
        </asp:Table>
        <br />
        <asp:Button ID="Button_OK" runat="server" OnClick="Button_OK_Click" Text="OK" Width="112px" />
        &nbsp;
    
    </div>
    </form>
</body>
</html>
