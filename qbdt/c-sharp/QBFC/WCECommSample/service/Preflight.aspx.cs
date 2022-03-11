using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Timers;

public partial class Preflight : System.Web.UI.Page
{
    //Private variables
    private string dbRelativePath = "\\db\\ecommdb.mdb";
    private string[,] customers;
    private string[,] sales;
    private string ticket;
    private static System.Timers.Timer aTimer;
    DataBaseManager dbm;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        ticket = Request["ticket"].ToString();
        Session["DBPath"] = HttpContext.Current.Server.MapPath("..") + dbRelativePath;
        dbm = new DataBaseManager((string)Session["DBpath"]);
        GetCustomers();
        DisplayCustomers();
        GetSalesReceipts();
        DisplaySalesReceipts();
        SetTimer();
    }

    void SetTimer() {
        aTimer = new System.Timers.Timer();
        aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
        aTimer.Interval = 10000; // Set the Interval to 10 seconds (10000 milliseconds).
        aTimer.Enabled = true;
    }
    
    void OnTimedEvent(object source, ElapsedEventArgs e) {
        //Enough time has passed but user didn't respond. 
        //Attempt to time out. 
        //User will be prompted.
        //closeBrowser();
    }

    protected void Button_OK_Click(object sender, EventArgs e) {
        dbm.SetInteractiveMode(ticket, "DONE");
        closeBrowser();
    }
    
    protected void Button_Cancel_Click(object sender, EventArgs e) {
        dbm.SetInteractiveMode(ticket, "CANCELLED");
        closeBrowser();
    }
    
    void closeBrowser() {
        this.Response.Write("<script language='javascript'>{ window.close(); }</script>");
    }

    void GetCustomers() {
        customers=dbm.GetAllNewCustomers();
    }
    
    void GetSalesReceipts() {
        sales = dbm.GetAllSalesReceipts();
    }
    
    void DisplayCustomers() {
        DisplayCustomerHeader();
        int length=customers.GetLength(0);
        for (int i = 0; i < length; i++) {
            TableRow r = new TableRow();

            string firstname    = (customers[i, 0] != null) ? customers[i, 0] : "-";
            string lastname     = (customers[i, 1] != null) ? customers[i, 1] : "-";
            string email        = (customers[i, 8] != null) ? customers[i, 8] : "-";

            CreateTableCell(r, firstname);
            CreateTableCell(r, lastname);
            CreateTableCell(r, email);

            Table_Customers.Rows.Add(r);
        }
    }
    
    void DisplaySalesReceipts() {
        DisplaySalesReceiptHeader();
        int length = sales.GetLength(0);
        for (int i = 0; i < length; i++) {
            TableRow r = new TableRow();

            string ordernumber      = (sales[i, 0] != null) ? sales[i, 0] : "-";
            string customerfullname = (sales[i, 1] != null) ? sales[i, 1] : "-";
            string itemfullname     = (sales[i, 2] != null) ? sales[i, 2] : "-";
            string itemrate         = (sales[i, 3] != null) ? sales[i, 3] : "-";
            string quantityordered  = (sales[i, 4] != null) ? sales[i, 4] : "-";

            CreateTableCell(r, ordernumber);
            CreateTableCell(r, customerfullname);
            CreateTableCell(r, itemfullname);
            CreateTableCell(r, itemrate);
            CreateTableCell(r, quantityordered);

            Table_SalesReceipts.Rows.Add(r);
        }
    }

    void DisplayCustomerHeader() {
        TableRow head = new TableRow();

        TableCell firstname_head = new TableCell();
        firstname_head.Controls.Add(new LiteralControl("First Name"));
        firstname_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(firstname_head);

        TableCell lastname_head = new TableCell();
        lastname_head.Controls.Add(new LiteralControl("Last Name"));
        lastname_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(lastname_head);

        TableCell email_head = new TableCell();
        email_head.Controls.Add(new LiteralControl("Email"));
        email_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(email_head);

        Table_Customers.Rows.Add(head);
    }
    
    void DisplaySalesReceiptHeader() {
        TableRow head = new TableRow();

        TableCell ordernumber_head = new TableCell();
        ordernumber_head.Controls.Add(new LiteralControl("Order Number"));
        ordernumber_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(ordernumber_head);

        TableCell customer_head= new TableCell();
        customer_head.Controls.Add(new LiteralControl("Customer Name"));
        customer_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(customer_head);

        TableCell itemname_head = new TableCell();
        itemname_head.Controls.Add(new LiteralControl("Item Name"));
        itemname_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(itemname_head);

        TableCell itemrate_head = new TableCell();
        itemrate_head.Controls.Add(new LiteralControl("Item Rate"));
        itemrate_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(itemrate_head);

        TableCell quantity_head = new TableCell();
        quantity_head.Controls.Add(new LiteralControl("Quantity Ordered"));
        quantity_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(quantity_head);

        Table_SalesReceipts.Rows.Add(head);
    }

    TableCell CreateTableCell(TableRow r, string cellname) {
        TableCell cell = createTableCellWithString(cellname);
        r.Cells.Add(cell);
        return cell;
    }
    
    TableCell createTableCellWithString(string textToShow) {
        TableCell cell = new TableCell();
        cell.Controls.Add(new LiteralControl(textToShow));
        cell.HorizontalAlign = HorizontalAlign.Center;
        return cell;
    }
}
