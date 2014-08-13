using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Summary description for Customer
/// Purpose of this class is to demonstrate how to create a customer object that will map 
/// to a client aka customer in access database. End user can sign up (creates a new customer) 
/// or login (with existing customer object).
/// 
/// Most of this class is created as an example, but code can be reused for real life customer 
/// scenario. The class name Customer should not be confused with a QuickBooks Customer. 
/// 
/// </summary>
public class Customer
{
    //Private Properties
    private string m_UserName;
    private string m_Password;
    // TODO: You should encrypt password before storing it.
    private string m_FirstName;
    private string m_LastName;
    private string m_Addr1;
    private string m_Addr2;
    private string m_City;
    private string m_State;
    private string m_Zip;
    private string m_Phone;
    private string m_Email;


    //Construct a customer object for a particular client
    public Customer(string p_UserName, string p_Password, string p_FirstName, string p_LastName)
	{
        //These three are required fields.
        //You can modify this per your requirement.
        m_UserName = p_UserName;
        m_Password = p_Password;
        m_FirstName = p_FirstName;
        m_LastName = p_LastName;
	}
    
    //Public methods
    public string UserName {
        get {
            return m_UserName;
        }
        set {
            m_UserName = value;
        }
    }
    public string Password{
        get {
            return m_Password;
        }
        set {
            m_Password = value;
        }
    }
    public string FirstName {
        get {
            return m_FirstName;
        }
        set {
            m_FirstName = value;
        }
    }
    public string LastName {
        get {
            return m_LastName;
        }
        set {
            m_LastName = value;
        }
    }
    public string Addr1{
        get {
            return m_Addr1;
        }
        set {
            m_Addr1 = value;
        }
    }
    public string Addr2 {
        get {
            return m_Addr2;
        }
        set {
            m_Addr2 = value;
        }
    }
    public string City{
        get {
            return m_City;
        }
        set {
            m_City= value;
        }
    }
    public string State {
        get {
            return m_State;
        }
        set {
            m_State= value;
        }
    }
    public string Zip{
        get {
            return m_Zip;
        }
        set {
            m_Zip = value;
        }
    }
    public string Phone{
        get {
            return m_Phone;
        }
        set {
            m_Phone = value;
        }
    }
    public string Email{
        get {
            return m_Email;
        }
        set {
            m_Email= value;
        }
    }
}


