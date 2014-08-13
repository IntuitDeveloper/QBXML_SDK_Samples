using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class _Default : System.Web.UI.Page 
{
    //Private variables
    private string dbRelativePath = "\\WCECommSample\\db\\ecommdb.mdb";
    private DataBaseManager dbm;

    //Protected methods
    protected void Page_Load(object sender, EventArgs e)
    {
        Session["DBPath"] = Server.MapPath(dbRelativePath);
    }
    
    protected void Login1_Authenticate(object sender, AuthenticateEventArgs e) {
        if (!IsValidUser(Login1.UserName, Login1.Password)) {
            Login1.InstructionText = "Enter a valid user.";
            Login1.InstructionTextStyle.ForeColor = System.Drawing.Color.RosyBrown;
            e.Authenticated = false;
        }
        else {
            Login1.InstructionText = "Login successful";
            e.Authenticated = true;
        }
        //Store successful authentication to a session variable for later check.
        Session["UserAuthenticated"] = e.Authenticated.ToString().ToUpper();
    }
    
    protected void Login1_LoggedIn(object sender, EventArgs e) {
        Session["Username"] = Login1.UserName;
        Server.Transfer("purchase.aspx");
    }

    
    //Private helper methods
    private void Debug(string text) {
        TextBox_Debug.Text = TextBox_Debug.Text + text + "\r\n";
    }
    
    bool IsValidUser(string username, string password) {
        dbm = new DataBaseManager((string)Session["DBpath"]);
        string[] login = dbm.GetCustomerLogin(username);
        if (password != login[1]) {
            return false;
        }
        return true;
    }
}
