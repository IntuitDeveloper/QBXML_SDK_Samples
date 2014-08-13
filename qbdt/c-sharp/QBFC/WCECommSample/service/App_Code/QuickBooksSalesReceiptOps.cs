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

/// <summary>
/// Operations with QuickBooks SalesReceipt objects.
/// </summary>
public class QuickBooksSalesReceiptOps
{
    //Constructor
    public QuickBooksSalesReceiptOps()
	{
	}

    public static String addSalesReceipts(Session sess, string[,] sales) {
        int length = sales.GetLength(0);
        if (length <= 0) {
            return ""; //No data to exchange. QBWC will call getLastError()
        }
        string salesReceiptAddRqXML = XmlManager.buildSalesReceiptAddRqXML(sess, sales);
        return salesReceiptAddRqXML;
    }

    public static string[,] extractSalesReceiptAddResponse(String response, Session sess) {
        string[,] sales;
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetResponse responseSet = sessionManager.ToMsgSetResponse(response, sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        int count = responseSet.ResponseList.Count;
        sales = new string[count, 2];
        sales = getSalesReceipts(responseSet, count);
        return sales;
    }

    private static string[,] getSalesReceipts(IMsgSetResponse responseSet, int count) {
        int row = 0;
        string[,] sales = new string[count, 2];
        for (int x = 0; x < count; x++) {
            IResponse iresponse = responseSet.ResponseList.GetAt(x);
            if (iresponse.StatusCode == 0) {
                ISalesReceiptRet salesRet = (ISalesReceiptRet)(iresponse.Detail) as ISalesReceiptRet;
                String lookie = salesRet.ToString();
                sales[row, 0] = salesRet.RefNumber.GetValue();
                sales[row, 1] = salesRet.TxnID.GetValue();
                row++;
            }
        }
        return sales;
    }
}
