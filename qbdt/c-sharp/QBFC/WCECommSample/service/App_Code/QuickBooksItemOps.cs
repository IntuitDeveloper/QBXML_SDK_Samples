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
/// Operations with QuickBooks Item objects.
/// </summary>
public class QuickBooksItemOps
{
    //A chunkSize defines the maximum number of records returned from QB
    private static int chunkSize = 100; 
    
    public QuickBooksItemOps()
	{
	}

    public static String queryAll(Session sess) {
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest(sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        requestSet.Attributes.OnError = ENRqOnError.roeStop;
        IItemInventoryQuery ItemQ = requestSet.AppendItemInventoryQueryRq();
        ItemQ.iterator.SetValue(ENiterator.itStart);
        ItemQ.ORListQuery.ListFilter.MaxReturned.SetValue(chunkSize);
        return requestSet.ToXMLString();
    }
    
    public static string[,] extractItemQueryResponses(String response, Session sess) {
        string[,] items;
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetResponse responseSet = sessionManager.ToMsgSetResponse(response, sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        IItemInventoryRetList itemRetList = responseSet.ResponseList.GetAt(0).Detail as IItemInventoryRetList;
        int length=itemRetList.Count;
        items=new string[length, 3];
        if (length > 0) {
            for (int ndx = 0; ndx < length; ndx++) {
                IItemInventoryRet itemRet = itemRetList.GetAt(ndx);
                items[ndx, 0] = itemRet.FullName.GetValue();
                items[ndx, 1] = itemRet.SalesPrice.GetValue().ToString();
                items[ndx, 2] = itemRet.QuantityOnHand.GetValue().ToString();
            }
        }
        return items;
    }

    public static Object hasMoreData(String response, Session sess) {
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetResponse responseSet = sessionManager.ToMsgSetResponse(response, sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        Object oReturn = null;
        int iRemain = responseSet.ResponseList.GetAt(0).iteratorRemainingCount;
        if (iRemain > 0) {
            oReturn = responseSet.ResponseList.GetAt(0).iteratorID; //generally a guid-like value 
        }
        return oReturn;
    }

    public static string queryNext(Session sess) {
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest(sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        requestSet.Attributes.OnError = ENRqOnError.roeStop;
        IItemInventoryQuery ItemQ = requestSet.AppendItemInventoryQueryRq();
        ItemQ.iterator.SetValue(ENiterator.itContinue);
        ItemQ.iteratorID.SetValue(sess.getProperty("queryContext").ToString());
        ItemQ.ORListQuery.ListFilter.MaxReturned.SetValue(chunkSize);
        return requestSet.ToXMLString();
    }
 
}

