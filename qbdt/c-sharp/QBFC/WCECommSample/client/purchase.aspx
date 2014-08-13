<%@ Page Language="C#" AutoEventWireup="true" CodeFile="purchase.aspx.cs" Inherits="purchase" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>WCECommSample</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <span><span style="font-family: Arial"><span style="font-size: 9pt"><span>Welcome</span>
        </span></span>
        </span>
        <asp:Label ID="Label_Username" runat="server" Width="169px" Font-Names="Arial" Font-Size="Small"></asp:Label><br />
        <span>
            <br />
            <span style="font-size: 9pt; font-family: Arial">
            Choose from list of available phone cards.</span></span><br />
        <asp:Table ID="Table1" runat="server" Width="533px" BorderStyle="Ridge" Font-Names="Arial" Font-Size="Small" GridLines="Both" BackColor="#FFFFC0">
        </asp:Table>
        <span style="font-size: 9pt; font-family: Arial">
        <br />
        </span>
        <asp:Button ID="submit_button" runat="server" Font-Names="Arial" Font-Size="Smaller"
            OnClick="submit_button_Click" Text="View shopping cart" /><br />
        <br />
        <span style="font-size: smaller; font-family: Arial">
        &nbsp;<br />
        <br />
        </span>
        <br />
        <br />
        <br />
        <br />
        <asp:TextBox ID="TextBox_Debug" runat="server" Font-Names="Courier New" Font-Size="Smaller"
            Height="283px" TextMode="MultiLine" Width="530px" Visible="False"></asp:TextBox></div>
    </form>
</body>
</html>
