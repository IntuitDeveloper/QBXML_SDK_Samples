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
/// Summary description for SampleAuthentication
/// </summary>
public class SampleAuthentication: IAuthentication
{
    private AccessDatabase db;
    
    //Constructor
    public SampleAuthentication() {
	}

    public void CreateAuthentication(Object context) {
        if (context != null) {
            db.setServer(context.ToString());
        }
    }

    public Object authenticate(String username, String password, Object context) {
        String ticket = null;
        db = (AccessDatabase)context;
        try {
            ticket = db.getTicket(username, password);
        }
        catch (AccessDatabaseError e) {
            throw new AuthenticateExceptionInvalidCredentials();
        }
        catch (Exception e2) {
            // A temporary exception, retry later
            throw new AuthenticateException(e2.ToString());
        }
        return ticket;
    }
}

