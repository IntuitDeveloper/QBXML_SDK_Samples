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

public partial class purchase : System.Web.UI.Page
{
    //Private variables
    DataBaseManager dbm;

    //Protected methods
    protected void Page_Load(object sender, EventArgs e)
    {
        Label_Username.Text = (string)Session["Username"];
        if(!Convert.ToBoolean((string)Session["UserAuthenticated"])){
            //This is to prevent user from by-passing the login by 
            //trying to go to purchase.aspx url.
            Server.Transfer("Default.aspx");
        }
        //User authenticated - proceed
        dbm = new DataBaseManager((string)Session["DBPath"]);
        //Load data from items table into Hashtable
        Hashtable items = dbm.GetItems();
        if (items.Count > 0) {
            //Present items to user
            DisplayItems(items);
            if ((string)Session["OrderNumber"] == null) {
                // Get a new order number for this purchase
                int ordernumber = dbm.GetNextOrderNumber();
                Session["OrderNumber"] = ordernumber.ToString();
            }
        }
    }
    
    protected void submit_button_Click(object sender, EventArgs e) {
        // Show the cart
        Server.Transfer("cart.aspx");
    }

    //Private methods
    TextBox createTextBox(string ID) {
        TextBox textBox_order = new TextBox();
        textBox_order.ID = ID;
        textBox_order.Width = 30;
        textBox_order.Style["text-align"] = "center";
        return textBox_order;
    }
    
    Button createButton(string textToShow, string commandArg) {
        Button mybutton = new Button();
        mybutton.Text = textToShow;
        mybutton.CommandArgument = commandArg;
        mybutton.Click += new EventHandler(this.addtocartbutton_Click);
        return mybutton;
    }
    
    TableCell createTableCellWithString(string textToShow) {
        TableCell cell = new TableCell();
        cell.Controls.Add(new LiteralControl(textToShow));
        cell.HorizontalAlign = HorizontalAlign.Center;
        return cell;
    }
    
    TableCell createTableCellWithTextBox(TextBox textBox) {
        TableCell cell = new TableCell();
        cell.Controls.Add(textBox);
        cell.HorizontalAlign = HorizontalAlign.Center;
        return cell;
    }
    
    TableCell createTableCellWithButton(Button mybutton) {
        TableCell cell = new TableCell();
        cell.Controls.Add(mybutton);
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
        order_head.Controls.Add(new LiteralControl("Order"));
        order_head.HorizontalAlign = HorizontalAlign.Center;
        head.Cells.Add(order_head);
        
        Table1.Rows.Add(head);
    }
    
    void DisplayItems(Hashtable items) {
        DisplayItemHeader();
        foreach (DictionaryEntry entry in items) {
            Item item=entry.Value as Item;
            TableRow r = new TableRow();
            
            TableCell name = createTableCellWithString(item.Name);
            r.Cells.Add(name);

            TableCell price = createTableCellWithString("$" + item.SalesPrice.ToString());
            r.Cells.Add(price);

            TextBox textBox_order = createTextBox(item.Name);
            TableCell order = createTableCellWithTextBox(textBox_order);
            r.Cells.Add(order);

            //Note: -
            //We are using CommandArgument as a vehicle to transfer TextBox.ID
            Button button_AddToCart = createButton("Add to cart", textBox_order.ID);
            TableCell addtocart = createTableCellWithButton(button_AddToCart);            
            r.Cells.Add(addtocart);

            Table1.Rows.Add(r);
        }
    }
    
    void addtocartbutton_Click(Object sender, EventArgs e) {
        // Capture the order into Hashtable
        Button clickedButton = (Button)sender;
        TextBox textBox_order=(TextBox)Table1.FindControl(clickedButton.CommandArgument);
        // Retrieve item id from mdb->items
        // Note here that the textbox was named after the item name.
        // See comments inside DisplayItems(Hashtable items) 
        // and createButton(string textToShow, string commandArg)
        string item_id = dbm.GetItemID(textBox_order.ID).ToString(); 
        // Retrieve customer id from mdb->customers
        string customer_id = dbm.GetCustomerID((string)Session["Username"]).ToString();
        // Store itemid, customerid, qtyordered into mdb->cart
        // In real application, you should also handle the case where item is out-of-stock
        string order_number = (string)Session["OrderNumber"];
        dbm.StoreOrderToCart(order_number, item_id, customer_id, textBox_order.Text);
    }
    
    void Debug(string text) {
        //This is purely for debuggging purpose. You can delete this if you want.
        TextBox_Debug.Text = TextBox_Debug.Text + text + "\r\n";
    }
}
