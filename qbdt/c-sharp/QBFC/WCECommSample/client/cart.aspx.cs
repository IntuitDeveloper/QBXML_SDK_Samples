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

public partial class cart : System.Web.UI.Page
{
    //Private variables
    DataBaseManager dbm;
    string customerName = "";

    //Protected methods
    protected void Page_Load(object sender, EventArgs e)
    {
        TextBox_confirm.Visible = false;
        if (IsPostBack) {
            ClearConfirm();
        }
        // Read mdb->cart - retrieve info using Session["OrderNumber"] 
        // and present it to user
        string orderNumber=(string)Session["OrderNumber"];
        if (orderNumber == null) {
            // No order number 
            // Generally happens if user tries to load this page directly.
            DisplayText(TextBox_Order, "left", "Cart is empty");
            Button_order.Visible = false;
            Button_EmptyCart.Visible = false;
            return;
        }
        Button_order.Visible = true;
        DisplayText(TextBox_Order, "left", "Order number = " + orderNumber);
        dbm = new DataBaseManager((string)Session["DBPath"]);
        string[,] orders=dbm.GetOrderInfo(orderNumber);
        DisplayItems(orders);
    }
    
    protected void Button_Order_Click(object sender, EventArgs e) {
        // For sample, we are faking a purchase completed here. 
        // In real application, you would have purchase buslogic here.

        // Insert this order into mdb->sales table
        dbm.StoreSales((string)Session["OrderNumber"]);

        // Delete order from mdb->cart table because order has been completed.
        dbm.DeleteOrder((string)Session["OrderNumber"]);

        // Show confirmation
        TextBox_confirm.Visible = true;
        string thankyou = "Purchase completed.\r\nDear " + customerName + " - Thank you for your business with us.";
        DisplayText(TextBox_confirm, "left", thankyou);
    }
    
    protected void Button_EmptyCart_Click(object sender, EventArgs e) {
        // Delete records from mdb-> cart where ordernumber=Session["OrderNumber"]
        dbm.DeleteOrder((string)Session["OrderNumber"]);

        // Go back to purchase page
        Server.Transfer("purchase.aspx");

    }

    
    //Private methods
    TableCell createTableCellWithString(string textToShow) {
        TableCell cell = new TableCell();
        cell.Controls.Add(new LiteralControl(textToShow));
        cell.HorizontalAlign = HorizontalAlign.Center;
        return cell;
    }
    
    void DisplayItemHeader() {
        TableRow head = new TableRow();

        TableCell name_head = new TableCell();
        name_head.Controls.Add(new LiteralControl("Card Name"));
        name_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(name_head);

        TableCell price_head = new TableCell();
        price_head.Controls.Add(new LiteralControl("Price"));
        price_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(price_head);

        TableCell order_head = new TableCell();
        order_head.Controls.Add(new LiteralControl("Quantity Ordered"));
        order_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(order_head);

        TableCell subtotal_head = new TableCell();
        subtotal_head.Controls.Add(new LiteralControl("Subtotal"));
        subtotal_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(subtotal_head);

        Table1.Rows.Add(head);
    }
    
    void DisplayItems(string[,] orders) {
        DisplayItemHeader();
        // Loop through each order and find more info on item and customer
        int rows = orders.Length / 3; //3=itemid, customerid, qtyordered
        decimal subtotal, grandtotal = 0;
        for (int i = 0; i < rows; i++) {
            TableRow r = new TableRow();
            
            Item item = dbm.GetItem(orders[i, 0]);
            customerName = dbm.GetCustomerName(orders[i, 1]);
            TableCell name = createTableCellWithString(item.Name);
            r.Cells.Add(name);
            TableCell price= createTableCellWithString(item.SalesPrice.ToString());
            r.Cells.Add(price);

            int qtyOrdered = Convert.ToInt32(orders[i, 2]);
            TableCell qty = createTableCellWithString(qtyOrdered.ToString());
            r.Cells.Add(qty);

            subtotal = item.SalesPrice * qtyOrdered;
            TableCell st= createTableCellWithString(subtotal.ToString());
            r.Cells.Add(st);

            grandtotal = grandtotal + subtotal;
            
            Table1.Rows.Add(r);
        }
        DisplayText(TextBox_GT, "right", "Grand total = " + grandtotal);
    }
    
    void DisplayText(TextBox textBox, string align, string text) {
        textBox.Style["text-align"] = align;
        textBox.Text = text;
    }
    
    void ClearConfirm() {
        DisplayText(TextBox_confirm, "left", "");
    }
    
    void Debug(string text) {
        TextBox_Debug.Text = TextBox_Debug.Text + text + "\r\n";
    }
}
