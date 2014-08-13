using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class signup : System.Web.UI.Page
{
    //Private variables
    DataBaseManager dbm;

    //Protected methods
    protected void Page_Load(object sender, EventArgs e)
    {
        dbm = new DataBaseManager((string)Session["DBPath"]);
    }
    
    protected void Button_Signup_Click(object sender, EventArgs e) {
        string username     =TextBox_Username.Text;
        string password     = TextBox_Password.Text;
        string firstname    =TextBox_Firstname.Text;
        string lastname     =TextBox_Lastname.Text;
        string addr1        =TextBox_Addr1.Text;
        string addr2        =TextBox_Addr2.Text;
        string city         =TextBox_City.Text;
        string state        =TextBox_State.Text;
        string zip          =TextBox_Zip.Text;
        string phone        =TextBox_Phone.Text;
        string email        =TextBox_Email.Text;
        Customer customer = new Customer(username, password, firstname, lastname);
        customer.Addr1 = addr1;
        customer.Addr2 = addr2;
        customer.City = city;
        customer.State = state;
        customer.Zip = zip;
        customer.Phone = phone;
        customer.Email = email;
        if (username == "" || password == "" || firstname == "" || lastname == "") {
            Label_Welcome.ForeColor = System.Drawing.Color.Red;
            Label_Welcome.Text = "Required fields need to be completed.";
            return;
        }
        if (dbm.StoreNewCustomer(customer)) {
            Label_Welcome.ForeColor = System.Drawing.Color.Green;
            Label_Welcome.Text = "Welcome " + customer.FirstName + " " + customer.LastName + ", you have been registered. Click 'Login' to go back to the login page.";
        }
        else {
            Label_Welcome.ForeColor = System.Drawing.Color.Red;
            Label_Welcome.Text = "Error occured during registration. Make sure required fields are completed.";
        }
    }
    
    protected void Button_Login_Click(object sender, EventArgs e) {
        Server.Transfer("Default.aspx");
    }
}
