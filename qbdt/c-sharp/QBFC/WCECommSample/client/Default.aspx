<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>WCECommSample - Purchase phone cards online</title>
</head>
<body style="font-size: 12pt">
    <form id="form1" runat="server">
    <div>
        <span><span style="font-family: Courier New" tabindex="2"><span style="color: #3366ff">
            <span style="font-size: 14pt">
            <strong style="font-size: smaller; font-family: Arial">Simple E-Commerce Application for Web Connector
                <br />
            (WCECommSample)</strong><br />
            </span>
            <br />
        </span>
            <span style="font-size: 9pt; color: #3366ff;" tabindex="3"><span style="font-family: Arial">
                This simulated ecommerce (e.g., Phone Calling Card) store provides items (e.g.,
                Prepaid 
                <br />
                Phone Calling Cards) to individuals and businesses.&nbsp; In order to proceed, you will need a
                <br />
                login (username and password).&nbsp;
                <br />
                <br />
                You can use the sample login (username="username" and password="password") or 
                <br />
                click sign-up to setup a new user login.<br />
            </span>
            </span>
        </span>
            <br />
            <asp:Login ID="Login1" runat="server" BackColor="#FFFFC0" BorderStyle="None" Font-Names="Arial"
                Font-Size="Small" LoginButtonText="Login" OnAuthenticate="Login1_Authenticate"
                OnLoggedIn="Login1_LoggedIn" TitleText="Enter login information" UserNameLabelText="Username:"
                UserNameRequiredErrorMessage="Username is required." Width="358px" CreateUserText="Sign up" CreateUserUrl="~/signup.aspx" TabIndex="1" DisplayRememberMe="False">
                <LoginButtonStyle Width="100px" Font-Names="Courier New" Height="25px" />
                <TitleTextStyle HorizontalAlign="Justify" VerticalAlign="Top" />
                <TextBoxStyle Font-Names="Courier New" />
                <InstructionTextStyle Font-Names="Courier New" HorizontalAlign="Left" />
            </asp:Login>
            <span style="font-family: Courier New">
            <br />
            <br />
            <span style="font-size: 9pt"></span></span></span>
    
    </div>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:TextBox ID="TextBox_Debug"  runat="server" Height="213px" TextMode="MultiLine"
            Visible="False" Width="350px"></asp:TextBox>
    </form>
</body>
</html>
