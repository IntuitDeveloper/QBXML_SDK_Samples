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
/// Summary description for XmlManager
/// </summary>
public class XmlManager
{
    //Constructor
	public XmlManager(){
	}

    public static string buildCustomerAddRqXML(Session sess, string[,] customers) {
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest(sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        requestSet.Attributes.OnError = ENRqOnError.roeStop;

        buildCustomerRequest(customers, ref requestSet);
        return requestSet.ToXMLString();
    }

    public static string buildSalesReceiptAddRqXML(Session sess, string[,] sales) {
        QBSessionManager sessionManager = new QBSessionManager();
        IMsgSetRequest requestSet = sessionManager.CreateMsgSetRequest(sess.getCountry(), sess.getMajorVers(), sess.getMinorVers());
        requestSet.Attributes.OnError = ENRqOnError.roeStop;

        buildSalesRequest(sales, requestSet);
        return requestSet.ToXMLString();
    }

    private static void buildCustomerRequest(string[,] customers, ref IMsgSetRequest requestSet) {
        //Walk customers array and build CustomerAdd xmls
        int length = customers.GetLength(0);
        for (int x = 0; x < length; x++) {
            string firstname = (customers[x, 0] != null) ? customers[x, 0] : "";
            string lastname = (customers[x, 1] != null) ? customers[x, 1] : "";
            string addr1 = (customers[x, 2] != null) ? customers[x, 2] : "";
            string addr2 = (customers[x, 3] != null) ? customers[x, 3] : "";
            string city = (customers[x, 4] != null) ? customers[x, 4] : "";
            string state = (customers[x, 5] != null) ? customers[x, 5] : "";
            string zip = (customers[x, 6] != null) ? customers[x, 6] : "";
            string phone = (customers[x, 7] != null) ? customers[x, 7] : "";
            string email = (customers[x, 8] != null) ? customers[x, 8] : "";

            ICustomerAdd CustAdd = requestSet.AppendCustomerAddRq();
            CustAdd.Name.SetValue(firstname + " " + lastname);
            CustAdd.FirstName.SetValue(firstname);
            CustAdd.LastName.SetValue(lastname);
            CustAdd.BillAddress.Addr1.SetValue(addr1);
            CustAdd.BillAddress.Addr2.SetValue(addr2);
            CustAdd.BillAddress.City.SetValue(city);
            CustAdd.BillAddress.State.SetValue(state);
            CustAdd.BillAddress.PostalCode.SetValue(zip);
            CustAdd.Phone.SetValue(phone);
            CustAdd.Email.SetValue(email);
        }
    }
    
    private static void buildSalesRequest(string[,] sales, IMsgSetRequest requestSet) {
        //Walk sales array and build SalesReceiptAdd xmls
        int length = sales.GetLength(0);
        string[,] items = new string[length, 3];
        int itemRow = 0;
        for (int x = 0; x < length; x++) {
            string orderNumber = (sales[x, 0] != null) ? sales[x, 0] : "";
            string customerFullName = (sales[x, 1] != null) ? sales[x, 1] : "";
            string itemFullName = (sales[x, 2] != null) ? sales[x, 2] : "";
            string itemRate = (sales[x, 3] != null) ? sales[x, 3] : "";
            string quantityOrdered = (sales[x, 4] != null) ? sales[x, 4] : "";

            items[itemRow, 0] = itemFullName;
            items[itemRow, 1] = itemRate;
            items[itemRow, 2] = quantityOrdered;

            //Peek into the next orderNumber.
            if (x + 1 < length && sales[x + 1, 0] == orderNumber) {
                itemRow++;
                continue;
            }
            string[,] lineItems = UtilityManager.CompactThisTwoDimensionalArray(items);

            ISalesReceiptAdd SalesReceiptAdd = requestSet.AppendSalesReceiptAddRq();
            SalesReceiptAdd.RefNumber.SetValue(orderNumber);
            SalesReceiptAdd.CustomerRef.FullName.SetValue(customerFullName);

            //Walk items array and build the SalesReceiptAddRq xml
            for (int r = 0; r < lineItems.GetLength(0); r++) {
                itemFullName = lineItems[r, 0];
                itemRate = lineItems[r, 1];
                quantityOrdered = lineItems[r, 2];
                ISalesReceiptLineAdd SalesReceiptLine = SalesReceiptAdd.ORSalesReceiptLineAddList.Append().SalesReceiptLineAdd;
                SalesReceiptLine.ItemRef.FullName.SetValue(itemFullName);
                SalesReceiptLine.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(itemRate));
                SalesReceiptLine.Quantity.SetValue(Convert.ToDouble(quantityOrdered));
            }
            lineItems = UtilityManager.ResetThisTwoDimensionalArray(lineItems);
            itemRow = 0;
        }
    }
}
