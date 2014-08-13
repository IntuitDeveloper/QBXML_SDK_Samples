using System;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections;

/// <summary>
/// Summary description for DataBaseManager
/// </summary>
public class DataBaseManager
{
    //Private variables
    private string m_ConnectionString= "";

    //Constructor
    public DataBaseManager(string dbPath) {
        m_ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" + dbPath;
    }

    //Public methods
    public string ConnectionString{
        get {
            return m_ConnectionString;
        }
    }
    
    //Item related
    public Hashtable GetItems() {
        Hashtable items=new Hashtable();
        DataRowCollection drc = GetDataRow("items", "select name, salesprice, quantityonhand from items");
        foreach (DataRow dr in drc) {
            //read drc and load each datarow into Item object
            Item item = GetItem(dr);
            items.Add(dr[0].ToString(), item);
        }
        return items;    
    }
    
    public int GetItemID(string itemName) {
        int id = 0;
        DataRowCollection drc = GetDataRow("items", "select id from items where name='" + itemName + "';");
        foreach (DataRow dr in drc) {
            id = (int)dr[0];
        }
        return id;
    }
    
    public Item GetItem(string id) {
        Item item = null;
        DataRowCollection drc = GetDataRow("items", "select name, salesprice, quantityonhand from items where id=" + id + ";");
        foreach (DataRow dr in drc) {
            item = GetItem(dr);
        }
        return item;
    }
    
    //Customer related
    public int GetCustomerID(string userName) {
        int id = 0;
        DataRowCollection drc = GetDataRow("customers", "select id from customers where username='" + userName + "';");
        foreach (DataRow dr in drc) {
            id = (int)dr[0];
        }
        return id;
    }
    
    public string GetCustomerName(string id) {
        string name="";
        DataRowCollection drc = GetDataRow("customers", "select firstname, lastname from customers where id=" + id + ";");
        foreach (DataRow dr in drc) {
            name = dr[0].ToString()+" "+dr[1].ToString();
        }
        return name;
    }
    
    public bool StoreOrderToCart(string orderNumber, string itemID, string customerID, string qtyOrdered) {
        int count = ExecuteDML("insert into cart (ordernumber, itemid, customerid, qtyordered) values (" + orderNumber + "," + itemID + "," + customerID + "," + qtyOrdered + ");");
        if (count > 0) {
            return true;
        }
        return false;
    }

    public int GetNextCustomerNumber() {
        return GetNextNumber("customers", "id");
    }

    public string[] GetCustomerLogin(string userName) {
        string[] str = new string[2];
        DataRowCollection dra = GetDataRow("customers", "select username, password from customers where username='" + userName + "';");
        foreach (DataRow dr in dra) {
            str[0] = dr[0].ToString(); //username
            str[1] = dr[1].ToString(); //password
        }
        if (dra.Count == 0) {
            str[0] = "No records found";
            return str;
        }
        else {
            return str;
        }
    }

    public bool StoreNewCustomer(Customer customer) {
        //Setting this as a NEW (=1 or true) customer. Once it is sent to QB, we will set this to OLD (=0 or false).
        int count = ExecuteDML("insert into customers values(" + GetNextCustomerNumber() + ",'" + customer.UserName + "','" + customer.Password + "','" + customer.FirstName + "','" + customer.LastName + "','" + customer.Addr1 + "','" + customer.Addr2 + "','" + customer.City + "','" + customer.State + "','" + customer.Zip + "','" + customer.Phone + "','" + customer.Email + "',1);");
        if (count > 0) {
            return true;
        }
        return false;
    }
    
    //Order related
    public string[,] GetOrderInfo(string orderNumber){
        DataRowCollection drc = GetDataRow("cart", "select itemid, customerid, qtyordered from cart where ordernumber="+orderNumber+";");
        string[,] orders = new string[drc.Count, 3]; 
        int row=0;
        foreach (DataRow dr in drc) {
            orders[row, 0]  =dr[0].ToString();  //itemid
            orders[row, 1]  = dr[1].ToString(); //customerid
            orders[row, 2]  = dr[2].ToString(); //qtyordered
            row++;
        }
        return orders;
    }
    
    public int GetNextOrderNumber() {
        return GetNextNumber("sales", "ordernumber");
    }
    
    public bool DeleteOrder(string orderNumber) {
        int count = ExecuteDML("delete from cart where ordernumber =" + orderNumber + ";");
        if (count > 0) {
            return true;
        }
        return false;
    }

    public bool StoreSales(string orderNumber) {
        int count = 0;
        string onum, custid, itemid, qtyordered;
        string[] order = new string[4]; //ordernumber, customerid, itemid, quantityordered
        DataRowCollection drc = GetDataRow("cart", "select ordernumber, customerid, itemid, qtyordered from cart where ordernumber=" + orderNumber);
        foreach (DataRow dr in drc) {
            onum        = dr[0].ToString(); //ordernumber
            custid      = dr[1].ToString(); //customerid
            itemid      = dr[2].ToString(); //itemid
            qtyordered  = dr[3].ToString(); //quantityordered
            count = count + ExecuteDML("insert into sales (ordernumber, customerid, itemid, quantityordered) values(" + onum + "," + custid + "," + itemid + "," + qtyordered + ");");
        }
        if (count > 0) {
            return true;
        }
        return false;
    }
    
    
    
    //Private methods
    private int GetNextNumber(string table, string field) {
        int n = 0;
        DataRowCollection dra = GetDataRow(table, "select max(" + field + ") from " + table + ";");
        foreach (DataRow dr in dra) {
            string temp = dr[0].ToString();
            if (temp != "") {
                n = Convert.ToInt32(temp);
            }
        }
        return n + 1;
    }
    
    private int ExecuteDML(string sqlString) {
        int count = 0;
        OleDbConnection myAccessConn = new OleDbConnection(m_ConnectionString);
        myAccessConn.Open();
        OleDbCommand myAccessCommand = new OleDbCommand(sqlString, myAccessConn);
        try {
            // count=no of rows effected
            count = myAccessCommand.ExecuteNonQuery();
        }
        finally {
            myAccessConn.Close();
        }
        return count;
    }
    
    private Item GetItem(DataRow dr) {
        if (dr == null) {
            return null;
        }
        Item item = new Item(dr[0].ToString());
        item.SalesPrice = Convert.ToDecimal(dr[1].ToString());
        item.QuantityOnHand = Convert.ToInt32(dr[2].ToString());
        return item;
    }
    
    private DataRowCollection GetDataRow(string tableName, string sqlString) {
        OleDbConnection myAccessConn = new OleDbConnection(m_ConnectionString);
        OleDbCommand myAccessCommand = new OleDbCommand(sqlString, myAccessConn);
        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
        myAccessConn.Open();
        DataSet myDataSet = new DataSet();
        try {
            myDataAdapter.Fill(myDataSet, tableName);
        }
        finally {
            myAccessConn.Close();
        }
        DataRowCollection drc = myDataSet.Tables[tableName].Rows;
        return drc;
    }

}
