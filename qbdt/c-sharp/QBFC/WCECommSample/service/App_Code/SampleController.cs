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
using Interop.QBFC13;

/// <summary>
/// SampleController implements the functionality of the original WCWebService
/// sample. I.e. it will do three send/receive cycles containing:
///     CustomerAdd
///     SalesReceiptAdd
///     ItemInventoryQuery
/// </summary>
public class SampleController: IController
{
    /// <summary>
    /// In general we transition from Do*Query to Get*Response,
    /// but if there is more data than will fit in one response,
    /// then we go into a More*Query state which alternates
    /// 
    /// </summary>
    enum ControllerStates {
        //Start = 0,
        Preflight=0,
        CustomerAddRq,
        CustomerAddRs,
        SalesReceiptAddRq,
        SalesReceiptAddRs,
        ItemInventoryQueryRq,
        MoreItemInventoryQueryRq,
        ItemInventoryQueryRs,
        Postflight,
        End = 9999
    };

    //Private variables
    private bool            dataToExchange = false;
    private string          dbRelativePath = "\\WCECommSample\\db\\ecommdb.mdb";
    private DataBaseManager dbm;
    
	public SampleController()
	{
        //Connect to ecommdb
        dbm = new DataBaseManager(HttpContext.Current.Server.MapPath(dbRelativePath));
    }

    /// <summary>
    /// Called by authenticate to determine if any data to exchange.
    /// </summary>
    /// <param name="sess"></param>
    /// <returns>true/false</returns>
    public bool haveAnyWork(Session sess)
    {
        //Count new customers in ecommdb 
        CountNewCustomers(sess);
        //Count new salesreceipts in ecommdb 
        CountNewSalesReceipts(sess);
        // Return status 
        return dataToExchange;
    }

    /// <summary>
    /// Called by sendRequestXML to determine what to send next.
    /// Corresponds to states Start and *Query states.
    /// </summary>
    /// <param name="sess"></param>
    /// <returns>String</returns>
    public String getNextAction(Session sess)
    {
        if (sess == null) throw new Exception("getNextAction: invalid session object");
        String action;
        
        ControllerStates controllerState = ControllerStates.Preflight;
        if (sess.getProperty("controllerState") != null){
            controllerState = (ControllerStates)sess.getProperty("controllerState");
        }
        ControllerStates lastControllerState = ControllerStates.Preflight;
        if (sess.getProperty("lastControllerState") != null) {
            lastControllerState = (ControllerStates)sess.getProperty("lastControllerState");
        }
        
        switch (controllerState) {
            case ControllerStates.Preflight:
                // Ask user if they want to do this update or not
                // if OK,       controllerState=ControllerStates.CustomerAddRq
                // if Cancel,   controllerState=ControllerStates.End
                dbm.SetInteractiveMode(sess.getTicket(), "NEEDED");
                action=""; //The empty string will trigger getLastError() which will start interactive mode.
                lastControllerState = ControllerStates.Preflight;
                controllerState = ControllerStates.CustomerAddRq;
                break;
            case ControllerStates.CustomerAddRq:
                // Build and send CustomerAddRq xml 
                string[,] customers = dbm.GetAllNewCustomers();
                action = QuickBooksCustomerOps.addCustomers(sess, customers);
                lastControllerState = ControllerStates.CustomerAddRq;
                controllerState = ControllerStates.CustomerAddRs;
                break;
            case ControllerStates.SalesReceiptAddRq:
                // Build and send SalesReceiptAddRq xml
                string[,] sales = dbm.GetAllSalesReceipts();
                action = QuickBooksSalesReceiptOps.addSalesReceipts(sess, sales);
                lastControllerState = ControllerStates.SalesReceiptAddRq;
                controllerState = ControllerStates.SalesReceiptAddRs;
                break;
            case ControllerStates.ItemInventoryQueryRq:
                // Build and send ItemInventoryQueryRq xml 
                action = QuickBooksItemOps.queryAll(sess);
                lastControllerState = ControllerStates.ItemInventoryQueryRq;
                controllerState = ControllerStates.ItemInventoryQueryRs;
                break;
            case ControllerStates.MoreItemInventoryQueryRq:
                // Build and send more ItemInventoryQueryRq xml 
                action = QuickBooksItemOps.queryNext(sess);
                lastControllerState = ControllerStates.MoreItemInventoryQueryRq;
                controllerState = ControllerStates.ItemInventoryQueryRs;
                break;
            case ControllerStates.Postflight:
                // Show user the data that was updated
                // Only OK,       controllerState=ControllerStates.End
                dbm.SetInteractiveMode(sess.getTicket(), "NEEDED");
                action = ""; //The empty string will trigger getLastError() which will start interactive mode.
                lastControllerState = ControllerStates.Postflight;
                controllerState = ControllerStates.End;
                break;
            case ControllerStates.End:
                // Send a dummy HostQuery so the control flows into processLastAction()
                lastControllerState = ControllerStates.End;
                action = "<?xml version=\"1.0\" ?><?qbxml version=\"2.0\"?>" + 
                    "<QBXML><QBXMLMsgsRq onError = \"stopOnError\"><HostQueryRq requestID = \"0\">" + 
                    "<IncludeRetElement>QBFileMode</IncludeRetElement></HostQueryRq></QBXMLMsgsRq></QBXML>";
                break;
            default:
                throw new Exception("getNextAction: Unexpected state: " + controllerState);
        }
        sess.setProperty("lastControllerState", lastControllerState);
        sess.setProperty("controllerState", controllerState);
        return action;
    }

    /// <summary>
    /// Called by receiveResponseXML to process the last data received.
    /// Corresponds to states Get*Response
    /// </summary>
    /// <param name="sess"></param>
    /// <param name="response"></param>
    /// <returns>int</returns>
    public int processLastAction(Session sess, String response)
    {
        int completion = 0;

        ControllerStates controllerState = ControllerStates.Preflight; 
        if (sess.getProperty("controllerState") != null)
        {
            controllerState = (ControllerStates)sess.getProperty("controllerState");
        }

        ControllerStates lastControllerState = ControllerStates.Preflight;
        if (sess.getProperty("lastControllerState") != null) {
            lastControllerState = (ControllerStates)sess.getProperty("lastControllerState");
        }

        switch (controllerState) {
            case ControllerStates.CustomerAddRs:
                //Parse response and store ListIDs in ecommdb->customers
                string[,] customers = QuickBooksCustomerOps.extractCustomerAddResponse(response, sess);
                dbm.SetListID(customers);
                lastControllerState = ControllerStates.CustomerAddRs;
                controllerState = ControllerStates.SalesReceiptAddRq;
                completion = 33;
                break;
            case ControllerStates.SalesReceiptAddRs:
                //Parse response and store TxnIDs in ecommdb->sales
                string[,] sales = QuickBooksSalesReceiptOps.extractSalesReceiptAddResponse(response, sess);
                dbm.SetTxnID(sales);
                lastControllerState = ControllerStates.SalesReceiptAddRs;
                controllerState = ControllerStates.ItemInventoryQueryRq;
                completion = 66;
                break;
            case ControllerStates.ItemInventoryQueryRs:
                //Parse response and store items in ecommdb->items
                string[,] items = QuickBooksItemOps.extractItemQueryResponses(response, sess);
                dbm.AddItems(items);
                Object queryContext = QuickBooksItemOps.hasMoreData(response, sess);
                if (queryContext != null) {
                    sess.setProperty("queryContext", queryContext);
                    lastControllerState = ControllerStates.ItemInventoryQueryRs;
                    controllerState = ControllerStates.MoreItemInventoryQueryRq;
                }
                else {
                    lastControllerState = ControllerStates.ItemInventoryQueryRs;
                    controllerState = ControllerStates.Postflight;
                    completion = 88;
                }
                break;
            case ControllerStates.End:
                //Finally, clean up work.
                //Following two lines are commented to be able to re-run the application.
                
                //For simplicty, we are assuming all customers made it to QB and therefore 
                //setting all customers to old. In real application, you would need to check
                //CustomerAddResponse and set them accordingly.
                //dbm.SetCustomerStatus(true);
                
                //Delete records from sales table. Note above applies to SalesReceipts as well.
                //dbm.DeleteAllSalesReceipts();
                
                completion = 100;
                break;
            default:
                throw new Exception("processLastAction: Unexpected state: " + controllerState);
        }

        sess.setProperty("lastControllerState", lastControllerState);
        sess.setProperty("controllerState", controllerState);
        return completion;
    }

    /// <summary>
    /// Called by isInteractiveDone() to check status of 
    /// ecommdb->interactive->interactiveMode 
    /// </summary>
    /// <param name="sess"></param>
    /// <returns>string</returns>
    public string getInteractive(Session sess) {
        string interactiveMode=dbm.GetInteractiveMode(sess.getTicket());
        return interactiveMode;
    }

    /// <summary>
    /// Called by getInteractiveURL() to store sessionid (for QB) into 
    /// ecommdb->interactive->sessionID
    /// </summary>
    /// <param name="sess"></param>
    /// <param name="sessionid"></param>
    /// <returns>void</returns>
    public void putInteractive(Session sess, string sessionid) {
        dbm.SetInteractiveSessionID(sess.getTicket(), sessionid);
        return;
    }

    /// <summary>
    /// Called by isInteractiveDone() to clear interactive session record from
    /// ecommdb->interactive table
    /// </summary>
    /// <param name="sess"></param>
    /// <returns>void</returns>
    public void removeInteractive(Session sess) {
        dbm.RemoveInteractive(sess.getTicket());
        return;
    }


    // PRIVATE METHODS
    private void CountNewCustomers(Session sess) {
        int length = dbm.CountCustomers();
        if (length > 0) dataToExchange = true;
    }
    private void CountNewSalesReceipts(Session sess) {
        int length = dbm.CountSalesReceipts();
        if (length > 0) dataToExchange = true;
    }

}
