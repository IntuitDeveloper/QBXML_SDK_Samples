<%@ Page Language="C#" AutoEventWireup="true" CodeFile="signup.aspx.cs" Inherits="signup" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Sign up</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <span style="font-family: Arial; font-size: smaller;">
            <asp:Label ID="Label_cart" runat="server" Font-Bold="True" Font-Names="Arial"
                Font-Size="X-Large" ForeColor="#8080FF" Text="New User Sign Up"></asp:Label><span
                    style="font-family: Arial">&nbsp;
            <br />
            <br />
            Enter the following information and click "Sign up". Fields with asterisk(*) are
            required fields.
            <br />
            <br />
                </span>
            <asp:Panel ID="Panel1" runat="server" BackColor="#FFFFC0" Height="249px" Width="550px">
                <br />
                Username*&nbsp; 
                <asp:TextBox ID="TextBox_Username" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox>&nbsp;<br />
                Password*&nbsp; &nbsp;<asp:TextBox ID="TextBox_Password" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    TextMode="Password" Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                <br />
                First name*&nbsp;
                <asp:TextBox ID="TextBox_Firstname" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                Last name*&nbsp;
                <asp:TextBox ID="TextBox_Lastname" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                <br />
                Address1 &nbsp; &nbsp;
                <asp:TextBox ID="TextBox_Addr1" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                Address2 &nbsp; &nbsp;
                <asp:TextBox ID="TextBox_Addr2" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                City &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                <asp:TextBox ID="TextBox_City" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                <br />
                State &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;<asp:TextBox ID="TextBox_State" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                Zip code &nbsp; &nbsp; &nbsp;
                <asp:TextBox ID="TextBox_Zip" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                <br />
                Phone &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
                <asp:TextBox ID="TextBox_Phone" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox><br />
                Email &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;<asp:TextBox ID="TextBox_Email" runat="server" Font-Names="Courier New" Font-Size="Smaller"
                    Width="450px" BorderStyle="Outset"></asp:TextBox></asp:Panel>
            <span style="font-family: Arial">
            <br />
            </span>
            <asp:Label ID="Label_Welcome" runat="server" Font-Names="Arial" Font-Size="Smaller"></asp:Label><br />
            <br />
            <asp:Button ID="Button_Signup" runat="server" Font-Names="Arial" Font-Size="Smaller"
                OnClick="Button_Signup_Click" Text="Sign up" Width="170px" /><span style="font-family: Arial">
                </span>
            <asp:Button ID="Button_Login" runat="server" Font-Names="Arial" Font-Size="Smaller"
                OnClick="Button_Login_Click" Text="Login" Width="170px" /><br />
        </span></div>
    </form>
</body>
</html>
