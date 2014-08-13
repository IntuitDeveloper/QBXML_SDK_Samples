using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections;

/// <summary>
/// Summary description for Session
/// </summary>
public class Session
{
    private String  ticket;
    private String  username;
    private String  password;
    private String  companyName ="";
    private String  country     ="US";
    private short   majorVers   =2;
    private short   minorVers   =0;
    private Hashtable sessionProperties = new Hashtable();

    public Session(String myTicket, String username, String password)
    {
        this.ticket = myTicket;
        this.username = username;
        this.password = password;
    }

	public void defineSession(String strCompanyFileName, String qbXMLCountry, short qbXMLMajorVers, short qbXMLMinorVers)
    {
        companyName = strCompanyFileName;
        country = qbXMLCountry;
        majorVers = qbXMLMajorVers;
        minorVers = qbXMLMinorVers;
    }

    public String getTicket()
    {
        return ticket;
    }

    public String getUsername()
    {
        return username;
    }

    public String getPassword()
    {
        return password;
    }

    public String getCompanyName()
    {
        return companyName;
    }

    public String getCountry()
    {
        return country;
    }

    public short getMajorVers()
    {
        return majorVers;
    }

    public short getMinorVers()
    {
        return minorVers;
    }

    public void setProperty(String name, Object value)
    {
        if (sessionProperties.Contains(name)) {
            sessionProperties[name] = value;
        }
        else {
            sessionProperties.Add(name, value);
        }
    }

    public Object getProperty(String name)
    {
        return (sessionProperties[name]);
    }
}
