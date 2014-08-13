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
/// ISessionPool is the interface for session classes. There can be many ways to
/// store information between calls from QBWC. Possibly implementations could
/// include memory sessions, database sessions, cookies, or whatever.
/// </summary>
public interface ISessionPool
{
    void put(String key, Session sess);
    Session get(String key);
    void invalidate(String key);
}
