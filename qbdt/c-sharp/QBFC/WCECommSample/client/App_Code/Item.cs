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
/// Summary description for Item
/// </summary>
public class Item
{
    //Private Properties
    private string  m_Name;
    private decimal m_SalesPrice;
    private int     m_QuantityOnHand;

    //Constructor
    public Item(string p_Name)
	{
        m_Name = p_Name;
	}

    //Public methods
    public string Name{
        get {
            return m_Name;
        }
        set {
            m_Name = value;
        }
    }
    public decimal SalesPrice{
        get {
            return m_SalesPrice;
        }
        set {
            m_SalesPrice = value;
        }
    }
    public int QuantityOnHand{
        get {
            return m_QuantityOnHand;
        }
        set {
            m_QuantityOnHand = value;
        }
    }
}
