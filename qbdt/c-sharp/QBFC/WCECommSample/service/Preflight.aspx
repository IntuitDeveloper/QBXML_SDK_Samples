<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Preflight.aspx.cs" Inherits="Preflight" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Pre-flight</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="font-size: smaller; font-family: Arial">
        <strong><span style="font-size: 16pt">PRE-FLIGHT INFORMATION<br />
        </span></strong>
        <br />
        The following customer and sales receipt data will be transferred to QuickBooks<br />
        <br />
        <asp:Label ID="Label1" runat="server" Text="New Customers"></asp:Label><br />
        <asp:Table ID="Table_Customers" runat="server" Width="593px" BorderStyle="Ridge" GridLines="Both">
        </asp:Table>
        &nbsp;&nbsp;<br />
        <asp:Label ID="Label2" runat="server" Text="New Sales Receipts"></asp:Label><br />
        <asp:Table ID="Table_SalesReceipts" runat="server" Width="591px" BorderStyle="Ridge" GridLines="Both">
        </asp:Table>
        <br />
        <asp:Button ID="Button_OK" runat="server" Text="OK" Width="112px" OnClick="Button_OK_Click" />
        &nbsp;
        <asp:Button ID="Button_Cancel" runat="server" Text="Cancel" Width="112px" OnClick="Button_Cancel_Click" /></div>
    </form>
</body>
</html>
