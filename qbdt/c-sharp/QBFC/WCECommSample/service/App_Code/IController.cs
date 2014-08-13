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
/// IController implements the behavior for this web service implementation.
/// </summary>
public interface IController
{
    bool    haveAnyWork(Session sess);
	String  getNextAction(Session sess);
    int     processLastAction(Session sess, String response);
    string  getInteractive(Session sess);
    void    putInteractive(Session sess, string sessionID);
    void    removeInteractive(Session sess);
}
