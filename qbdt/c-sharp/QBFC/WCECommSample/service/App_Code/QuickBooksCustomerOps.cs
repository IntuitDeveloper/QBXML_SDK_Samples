using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using Interop.QBFC13;

public class QuickBooksCustomerOps
{
    //Constructor
	public QuickBooksCustomerOps()
	{
    }

    public static String addCustomers(Session sess, string[,] customers) {
        int length = customers.GetLength(0);
        if (length <= 0) {
            return ""; 
        }
        string customerAddRqXML = XmlManager.buildCustomerAddRqXML(sess, customers);
        return customerAddRqXML;
    }
    
    public static string[,] extractCustomerAddResponse(String response, Session sess) {
        string[,] customers; 
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetResponse responseSet = sessionManager.ToMsgSetResponse(response, sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        int count = responseSet.ResponseList.Count;
        customers = new string[count, 3];
        customers=getCustomers(responseSet, count);
        return customers;
    }

    private static string[,] getCustomers(IMsgSetResponse responseSet, int count) {
        int row = 0;
        string[,] customers = new string[count, 3]; 
        for (int x = 0; x < count; x++) {
            IResponse iresponse = responseSet.ResponseList.GetAt(x);
            if (iresponse.StatusCode == 0) {
                ICustomerRet customerRet = (ICustomerRet)(iresponse.Detail) as ICustomerRet;
                String lookie = customerRet.ToString();
                customers[row, 0] = customerRet.FirstName.GetValue();
                customers[row, 1] = customerRet.LastName.GetValue();
                customers[row, 2] = customerRet.ListID.GetValue();
                row++;
            }
        }
        return customers;
    }
}
