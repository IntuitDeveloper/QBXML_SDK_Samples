<%@ Page Language="C#" AutoEventWireup="true" CodeFile="cart.aspx.cs" Inherits="cart" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Shopping cart</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="font-size: smaller; font-family: Arial">
        <asp:Label ID="Label_cart" runat="server" Font-Names="Arial" Font-Size="X-Large"
            Text="Shopping cart" Font-Bold="True"></asp:Label><br />
        <br />
        <asp:TextBox ID="TextBox_Order" runat="server" BorderStyle="None" Font-Names="Courier New"
            Font-Size="Smaller" ReadOnly="True" Width="525px"></asp:TextBox><br />
        <asp:Table ID="Table1" runat="server" BackColor="#FFFFC0" BorderStyle="Ridge" Font-Names="Arial"
            Font-Size="Small" GridLines="Both" Width="533px">
        </asp:Table>
        <asp:TextBox ID="TextBox_GT" runat="server" BorderStyle="None" Font-Names="Courier New"
            Font-Size="Smaller" ReadOnly="True" Width="525px"></asp:TextBox><br />
        <br /><asp:Button ID="Button_order" runat="server" Font-Names="Arial" Font-Size="10pt"
            OnClick="Button_Order_Click" Text="Check out" Width="210px" />
        <asp:Button ID="Button_EmptyCart" runat="server" Font-Names="Arial" Font-Size="10pt"
            OnClick="Button_EmptyCart_Click" Text="Empty cart" Width="210px" /><br />
        <br />
        <asp:TextBox ID="TextBox_confirm" runat="server" BackColor="White" BorderStyle="Dotted"
            Font-Names="Arial" Font-Size="10pt" Height="53px" TextMode="MultiLine"
            Width="531px" ReadOnly="True"></asp:TextBox>&nbsp;&nbsp;<br />
        <br />
        <br />
        <br />
        <asp:TextBox ID="TextBox_Debug" runat="server" Font-Names="Arial" Font-Size="Smaller"
            Height="306px" TextMode="MultiLine" Visible="False" Width="523px"></asp:TextBox></div>
    </form>
</body>
</html>
