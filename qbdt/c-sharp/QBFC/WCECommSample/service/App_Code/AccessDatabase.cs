using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

//using System.Xml;
//using System.Net;
//using System.Text;
//using System.IO;

/// <summary>
/// Summary description for AccessDatabaseError
/// </summary>
class AccessDatabaseError : Exception {
    public AccessDatabaseError(String reason) { }
}

/// <summary>
/// Summary description for AccessDatabase
/// For the purpose of sample, we are using 
/// simple access database based credential.
/// </summary>
public class AccessDatabase
{
    private string server = "\\WCECommSample\\db\\ecommdb.mdb";
    private string ticket = null;
    private DataBaseManager dbm;


    //Constructor
    public AccessDatabase(){
        if (dbm==null) dbm = new DataBaseManager(HttpContext.Current.Server.MapPath(server));
    }

    public AccessDatabase(String myTicket){
        ticket = myTicket;
    }

    
    //Public methods
    /// <summary>
    /// Set server name to build connection string for AccessDatabase.
    /// </summary>
    /// <param name="username"></param>
    /// <param name="password"></param>
    /// <returns>void</returns>
    public void setServer(String server) {
        this.server = server;
        //Connect to ecommdb
        dbm = new DataBaseManager(HttpContext.Current.Server.MapPath(server));
    }

    /// <summary>
    /// Get a new ticket for the indicated user and password.
    /// </summary>
    /// <param name="username"></param>
    /// <param name="password"></param>
    /// <returns>string</returns>
    public string getTicket(string username, string password) {
        if (ticket != null) {
            return ticket;
        }
        if (!dbm.ValidUser(username, password)) {
            return null;
        }
        ticket=dbm.GetTicket(username);
        if (ticket == null || ticket== "") {
            //Create one and store it in database.
            ticket = System.Guid.NewGuid().ToString();
            dbm.SetTicket(username, ticket);
        }
        return ticket;
    }

    /// <summary>
    /// If you authenticate with a separate AccessDatabase object you
    /// can use this to reuse the ticket you got upon authentication.
    /// </summary>
    /// <param name="myTicket"></param>
    public void setTicket(String myTicket) {
        ticket = myTicket;
    }

    /// <summary>
    /// Get the current ticket.
    /// </summary>
    /// <returns></returns>
    public string getTicket() {
        return (ticket);
    }
}
