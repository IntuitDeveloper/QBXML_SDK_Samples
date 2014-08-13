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
/// This class provides all the functions required to manage the back-end database (in this case, microsoft
/// access database). 
/// The connection string was one of the sample simplifications, and in real code you would need to have your 
/// own connection string pointing to your database. All the SQL used in this class are also access oriented. 
/// Therefore, if you are to use another database, use proper SQL syntax accordingly.
/// </summary>
public class DataBaseManager
{
    //Private variables
    private string m_ConnectionString= "";
    private const int CUSTOMER_TABLE_FIELD_COUNT=9;
    private const int SALES_TABLE_FIELD_COUNT=5;

    //Constructor
    public DataBaseManager(string dbPath) {
        m_ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" + dbPath;
    }

    #region Public methods
    public string ConnectionString{
        get {
            return m_ConnectionString;
        }
    }

    //Customers related
    public string GetListID(string firstname, string lastname) {
        //FirstName + Lastname not necessarily unique name in QuickBooks. We are using this just for 
        //sample purposes, but this is NOT what should be done in production. 
        string listid = "";
        DataRowCollection drc = GetDataRow("customers", "select listid from customers where firstname='" + firstname + "' and lastname='" + lastname + "';");
        foreach (DataRow dr in drc) {
            listid = dr[0].ToString();
        }
        return listid;
    }
    public string[,] GetAllNewCustomers() {
        DataRowCollection drc = GetDataRow("customers", "select firstname, lastname, addr1, addr2, city, state, zip, phone, email from customers where new=True;");
        //customer[row, 0]=firstname
        //customer[row, 1]=lastName
        //customer[row, 2]=addr1
        //customer[row, 3]=addr2
        //customer[row, 4]=city
        //customer[row, 5]=state
        //customer[row, 6]=zip
        //customer[row, 7]=phone 
        //customer[row, 8]=email
        string[,] customer = new string[drc.Count, CUSTOMER_TABLE_FIELD_COUNT];
        int row = 0;
        foreach (DataRow dr in drc) {
            GetCustomerRecord(customer, row, dr);
            row++;
        }
        return customer;
    }
    public void SetListID(string[,] customers) {
        for (int row = 0; row < customers.GetLength(0); row++) {
            ExecuteDML("update customers set listid='" + customers[row, 2] + "' where firstname='" + customers[row, 0] + "' and lastname='" + customers[row, 1] + "';");
        }
    }
    public bool SetCustomerStatus(bool toOld) {
        //This method is used to reset customer status.
        //For simplicty, we are assuming all customers made it to QB and therefore 
        //setting all customers to old. In real application, you would need to check
        //CustomerAddResponse and set them accordingly.
        string strNew = "True";
        if (toOld) {
            strNew = "False";
        }
        int count = ExecuteDML("update customers set new=" + strNew);
        if (count > 0) {
            return true;
        }
        return false;
    }
    public int CountCustomers() {
        return CountRows("customers", null);
    }

    //SalesReceipts related
    public string GetTxnID(string ordernumber) {
        string txnid = "";
        DataRowCollection drc = GetDataRow("sales", "select txnid from sales where ordernumber=" + ordernumber+ ";");
        foreach (DataRow dr in drc) {
            txnid = dr[0].ToString();
        }
        return txnid;
    }
    public string[,] GetAllSalesReceipts() {
        DataRowCollection drc = GetDataRow("sales", "select ordernumber, customerid, itemid, quantityordered from sales;");
        //sales[row, 0]=orderNumber
        //sales[row, 1]=customerFullName
        //sales[row, 2]=itemFullName
        //sales[row, 3]=itemRate
        //sales[row, 4]=quantityOrdered
        string[,] sales = new string[drc.Count, SALES_TABLE_FIELD_COUNT];
        int row = 0;
        foreach (DataRow dr in drc) {
            GetSalesRecord(sales, row, dr);
            row++;
        }
        return sales;
    }
    public void SetTxnID(string[,] sales) {
        for (int row = 0; row < sales.GetLength(0); row++) {
            ExecuteDML("update sales set txnid='" + sales[row, 1] + "' where ordernumber=" + sales[row, 0] + ";");
        }
    }
    public bool DeleteAllSalesReceipts() {
        int count = ExecuteDML("delete from sales;");
        if (count > 0) {
            return true;
        }
        return false;
    }
    public int CountSalesReceipts() {
        return CountRows("sales", null);
    }
    
    //Items related
    public void AddItems(string[,] items) {
        //Walk the items array and insert each record into ecommdb->items.
        int length=items.GetLength(0);
        if (length <= 0) {
            return;
        }
        for (int row = 0; row < length; row++) {
            string name             = items[row, 0];
            string salesprice       = items[row, 1];
            string quantityonhand   = items[row, 2];
            if (ItemExists(name)) {
                //if item already exists, update it.
                ExecuteDML("update items set salesprice= " + salesprice + ", quantityonhand=" + quantityonhand + " where name='" + name + "';");
            }
            else {
                //if item doesn't exists, insert it.
                ExecuteDML("insert into items (id, name, salesprice, quantityonhand) values(" + row + ",'" + name + "'," + salesprice + "," + quantityonhand + ");");
            }
        }
    }
    public bool ItemExists(string itemName) {
        int count = CountRows("items", "where name = '" + itemName + "';");
        if (count > 0) {
            return true;
        }
        return false;
    }

    //Interactive
    public string GetInteractiveMode(string ticket) {
        string interactiveMode = "";
        DataRowCollection drc = GetDataRow("interactive", "select interactivemode from interactive where ticket='" + ticket + "';");
        foreach (DataRow dr in drc) {
            interactiveMode = dr[0].ToString();
        }
        return interactiveMode;
    }
    public void SetInteractiveMode(string ticket, string mode) {
        int count = ExecuteDML("update interactive set interactivemode ='" + mode + "' where ticket = '" + ticket + "';");
        if (count <= 0) {
            count = ExecuteDML("insert into interactive(ticket, interactivemode) values('" + ticket + "','" + mode + "');");
        }
    }
    public void SetInteractiveSessionID(string ticket, string sessionID) {
        int count = ExecuteDML("update interactive set sessionid ='" + sessionID + "' where ticket = '" + ticket + "';");
        if (count <= 0) {
            count = ExecuteDML("insert into interactive(ticket, sessionid) values('" + ticket + "','" + sessionID + "');");
        }
    }
    public void RemoveInteractive(string ticket) {
        ExecuteDML("delete from interactive where ticket='" + ticket + "';");
    }
    #endregion

    //Credentials
    public bool ValidUser(string username, string password) {
        bool retVal = false;
        string pwd = "";
        DataRowCollection drc = GetDataRow("credentials", "select password from credentials where username='" + username + "';");
        foreach (DataRow dr in drc) {
            pwd = dr[0].ToString();
            if (pwd == password) {
                retVal = true;
                break;
            }
        }
        return retVal;
    }
    public string GetTicket(string username) {
        string ticket = "";
        DataRowCollection drc = GetDataRow("credentials", "select ticket from credentials where username='" + username + "';");
        foreach (DataRow dr in drc) {
            ticket= dr[0].ToString();
        }
        return ticket;
    }
    public void SetTicket(string username, string ticket) {
        int count = ExecuteDML("update credentials set ticket='" + ticket + "' where username = '" + username+ "';");
    }

    #region Private methods
    //Customers related
    private void GetCustomerRecord(string[,] customer, int row, DataRow dr) {
        //firstname -> Customer.FirstName
        customer[row, 0] = dr[0].ToString();

        //lastname -> Customer.LastName 
        customer[row, 1] = dr[1].ToString();

        //addr1 -> Customer.BillAddress.Addr1
        customer[row, 2] = dr[2].ToString();

        //addr2 -> Customer.BillAddress.Addr2
        customer[row, 3] = dr[3].ToString();

        //city -> Customer.BillAddress.City
        customer[row, 4] = dr[4].ToString();

        //state -> Customer.BillAddress.State
        customer[row, 5] = dr[5].ToString();

        //zip -> Customer.BillAddress.Zip
        customer[row, 6] = dr[6].ToString();
    }
    private string GetCustomerName(string id) {
        string name = "";
        DataRowCollection drc = GetDataRow("customers", "select firstname, lastname from customers where id=" + id + ";");
        foreach (DataRow dr in drc) {
            name = dr[0].ToString() + " " + dr[1].ToString();
        }
        return name;
    }

    //SalesReceipts related
    private void GetSalesRecord(string[,] sales, int row, DataRow dr) {
        //ordernumber -> RefNumber
        sales[row, 0] = dr[0].ToString();

        //customerid -> CustomerRef.FullName -> customers.firstname + " " + customers.lastname 
        sales[row, 1] = GetCustomerName(dr[1].ToString());

        //itemid -> Item.FullName, Item.Rate -> items.name, items.salesprice
        string[] item = GetItem(dr[2].ToString());
        sales[row, 2] = item[0].ToString();
        sales[row, 3] = item[1].ToString();

        //quantityordered -> Item.Quantity
        sales[row, 4] = dr[3].ToString();
    }
    private string[] GetItem(string id) {
        string[] item = new string[2];
        DataRowCollection drc = GetDataRow("items", "select name, salesprice from items where id=" + id + ";");
        foreach (DataRow dr in drc) {
            item[0] = dr[0].ToString(); //name
            item[1] = dr[1].ToString(); //salesprice
        }
        return item;
    }

    //Helpers
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
        //As the method name implies, this method covers all DML (Data Manipulation Language) 
        //calls i.e., Insert, Update, and Delete. 
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
    private int CountRows(string tableName, string condition) {
        int count = 0;
        string sqlString = "select count(*) from " + tableName;
        if (condition != null) sqlString = sqlString + " " + condition;
        OleDbConnection myAccessConn = new OleDbConnection(m_ConnectionString);
        OleDbCommand myAccessCommand = new OleDbCommand(sqlString, myAccessConn);
        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);
        myAccessConn.Open();
        try {
            count = Convert.ToInt32(myAccessCommand.ExecuteScalar().ToString());
        }
        finally {
            myAccessConn.Close();
        }
        return count;
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
    #endregion
}
