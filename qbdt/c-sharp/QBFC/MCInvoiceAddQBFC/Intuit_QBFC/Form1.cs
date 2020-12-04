// Copyright (c) 2008-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Interop.QBFC13;
using MCInvoiceAddQBFC.Session_Framework;


namespace MCInvoiceAddQBFC
{
    public partial class Form1 : Form
    {
        string workCurrency = "";
        string homeCurrency = "";
        double exchangeRate = 1.0;

        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // It is imperative that the SessionManager object be disposed of when the 
            // applicaiton closes!  If this is not done, there is a possibility that a
            // connection to QuickBooks will remain open.  This will preclude QuickBooks
            // from being able to close; the user will need to go into the Task Manager
            // and manually "kill" the QuickBooks application.
            SessionManager.getInstance().Dispose();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Show();
            showStatus("Checking Multi Currency support...");
            if (!multiCurrencyOn())
            {
                string msg = "Multi currency feature is not enabled \r\n" +
                    "for this company file. Use Edit -> Preferences -> \r\n" +
                    "Multiple Currencies to enable MultiCurrency \r\n" +
                    "support and then restart this sample.\r\n";
                MessageBox.Show(this, msg, "Exiting...");
                this.Close();
                return;
            }
            showStatus("Loading customers...");
            loadCustomers();
            showStatus("Loading items...");
            loadItems();
            showStatus("Loading terms...");
            loadTerms();
            showStatus("Loading sales tax codes...");
            loadSalesTaxCodes();
            showStatus("Loading customer msgs...");
            loadCustomerMsg();
            showStatus("Ready");
        }


        SessionManager sessionManager;
        private short maxVersion;


        private bool multiCurrencyOn()
        {
            bool ret = false;
            connectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildPreferencesQueryRq(new string[] { "MultiCurrencyPreferences" }, null));
            string[] prefs = parsePreferencesQueryRs(responseMsgSet, 2);
            ret = Convert.ToBoolean(prefs[0]);
            homeCurrency = prefs[1];
            disconnectFromQB();
            return ret;
        }
        private void loadCustomers()
        {
            string request = "CustomerQueryRq";
            connectToQB();
            int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerQueryRq(new string[] { "FullName" }, null));
            string[] customerList = parseCustomerQueryRs(responseMsgSet, count);
            disconnectFromQB();
            fillComboBox(this.comboBox_Customer, customerList);
        }
        private void loadItems()
        {
            string request = "ItemQueryRq";
            connectToQB();
            int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemQueryRq(new string[] { "FullName" }, null));
            string[,] retVal = parseItemQueryRs(responseMsgSet, count, 1);
            disconnectFromQB();
            string[] itemList = getItemFullNames(retVal);
            fillComboBox(comboBox_Item1, itemList);
            fillComboBox(comboBox_Item2, itemList);
            fillComboBox(comboBox_Item3, itemList);
            fillComboBox(comboBox_Item4, itemList);
            fillComboBox(comboBox_Item5, itemList);
        }
        private string[] getItemFullNames(string[,] retVal)
        {
            //retVal[countOfRows, arraySize]
            //arraySize is 3 for FullName, Desc, Price but in this case 1 for FullName only
            int countOfRows=retVal.GetUpperBound(0);
            string[] itemList=new string[countOfRows];
            for(int i=0; i<countOfRows; i++){
                itemList[i]=retVal[i, 0];
            }
            return itemList;
        }
        private void loadTerms()
        {
            string request = "TermsQueryRq";
            connectToQB();
            int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildTermsQueryRq());
            string[] termsList = parseTermsQueryRs(responseMsgSet, count);
            disconnectFromQB();
            fillComboBox(this.comboBox_Terms, termsList);
        }
        private void loadSalesTaxCodes()
        {
            string request = "SalesTaxCodeQueryRq";
            connectToQB();
            int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildSalesTaxCodeQueryRq());
            string[] salesTaxCodeList = parseSalesTaxCodeQueryRs(responseMsgSet, count);
            disconnectFromQB();
            fillComboBox(this.comboBox_Tax1, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax2, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax3, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax4, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax5, salesTaxCodeList);
        }
        private void loadCustomerMsg()
        {
            string request = "CustomerMsgQueryRq";
            connectToQB();
            int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerMsgQueryRq(new string[] { "Name" }, null));
            string[] customerMsgList = parseCustomerMsgQueryRs(responseMsgSet, count);
            disconnectFromQB();
            fillComboBox(comboBox_CustomerMessage, customerMsgList);
        }
        private string getBillShipTo(string customerName, string billOrShip)
        {
            connectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerQueryRq(new string[] { billOrShip }, customerName));
            string[] billShipTo = parseCustomerQueryRs(responseMsgSet, 1);
            if (billShipTo[0] == null) billShipTo[0] = "";
            disconnectFromQB();
            return billShipTo[0];
        }
        private string getCurrencyCode(string customerName)
        {
            connectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerQueryRq(new string[] { "CurrencyRef" }, customerName));
            string[] currencyCode = parseCustomerQueryRs(responseMsgSet, 1);
            disconnectFromQB();
            return currencyCode[0];
        }
        private string getExchangeRate(string currencyName)
        {
            connectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCurrencyQueryRq(currencyName));
            string[] exrate = parseCurrencyQueryRs(responseMsgSet, 1);
            disconnectFromQB();
            if (exrate[0] == null || exrate[0] == "") exrate[0] = "1.0";
            return exrate[0];
        }
        private void getItem(int index)
        {
            string[,] itemInfo;
            switch (index)
            {
                case 1:
                    itemInfo = getItemInfo(comboBox_Item1.Text);
                    textBox_Desc1.Text = itemInfo[0, 0];
                    textBox_Price1.Text = itemInfo[0, 1];
                    break;
                case 2:
                    itemInfo = getItemInfo(comboBox_Item2.Text);
                    textBox_Desc2.Text = itemInfo[0, 0];
                    textBox_Price2.Text = itemInfo[0, 1];
                    break;
                case 3:
                    itemInfo = getItemInfo(comboBox_Item3.Text);
                    textBox_Desc3.Text = itemInfo[0, 0];
                    textBox_Price3.Text = itemInfo[0, 1];
                    break;
                case 4:
                    itemInfo = getItemInfo(comboBox_Item4.Text);
                    textBox_Desc4.Text = itemInfo[0, 0];
                    textBox_Price4.Text = itemInfo[0, 1];
                    break;
                case 5:
                    itemInfo = getItemInfo(comboBox_Item5.Text);
                    textBox_Desc5.Text = itemInfo[0, 0];
                    textBox_Price5.Text = itemInfo[0, 1];
                    break;
            }
            return;
        }
        private string[,] getItemInfo(string itemName)
        {
            connectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemQueryRq(new string[] { "SalesOrPurchase" }, itemName));
            string[,] itemInfo = parseItemQueryRs(responseMsgSet, 1, 2);
            disconnectFromQB();
            return itemInfo;
        }


        // REQUEST BUILDING
        private IMsgSetRequest buildPreferencesQueryRq(string[] includeRetElement, string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest(); 
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IPreferencesQuery prefQuery = requestMsgSet.AppendPreferencesQueryRq();
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                prefQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }
        private IMsgSetRequest buildDataCountQuery(string request)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            switch (request)
            {
                case "CustomerQueryRq":
                    ICustomerQuery custQuery = requestMsgSet.AppendCustomerQueryRq();
                    custQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
                    break;
                case "ItemQueryRq":
                    IItemQuery itemQuery = requestMsgSet.AppendItemQueryRq();
                    itemQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
                    break;
                case "TermsQueryRq":
                    ITermsQuery termsQuery = requestMsgSet.AppendTermsQueryRq();
                    termsQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
                    break;
                case "SalesTaxCodeQueryRq":
                    ISalesTaxCodeQuery salesTaxQuery = requestMsgSet.AppendSalesTaxCodeQueryRq();
                    salesTaxQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
                    break;
                case "CustomerMsgQueryRq":
                    ICustomerMsgQuery custMsgQuery = requestMsgSet.AppendCustomerMsgQueryRq();
                    custMsgQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
                    break;
                default:
                    break;
            }
            return requestMsgSet;            
        }
        private IMsgSetRequest buildCustomerQueryRq(string[] includeRetElement, string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICustomerQuery custQuery = requestMsgSet.AppendCustomerQueryRq();
            if (fullName != null)
            {
                custQuery.ORCustomerListQuery.FullNameList.Add(fullName);
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                custQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }
        private IMsgSetRequest buildItemQueryRq(string[] includeRetElement, string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IItemQuery itemQuery = requestMsgSet.AppendItemQueryRq();
            if (fullName != null)
            {
                itemQuery.ORListQuery.FullNameList.Add(fullName);
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                itemQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }
        private IMsgSetRequest buildTermsQueryRq()
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ITermsQuery termsQuery = requestMsgSet.AppendTermsQueryRq();
            termsQuery.IncludeRetElementList.Add("Name");
            return requestMsgSet;
        }
        private IMsgSetRequest buildSalesTaxCodeQueryRq()
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ISalesTaxCodeQuery salesTaxCodeQuery = requestMsgSet.AppendSalesTaxCodeQueryRq();
            salesTaxCodeQuery.IncludeRetElementList.Add("Name");
            return requestMsgSet;
        }
        private IMsgSetRequest buildCustomerMsgQueryRq(string[] includeRetElement, string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICustomerMsgQuery custMsgQuery = requestMsgSet.AppendCustomerMsgQueryRq();
            if (fullName != null)
            {
                custMsgQuery.ORListQuery.FullNameList.Add(fullName);
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                custMsgQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }
        private IMsgSetRequest buildCurrencyQueryRq(string fullName)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICurrencyQuery currQuery = requestMsgSet.AppendCurrencyQueryRq();
            if (fullName != null)
            {
                currQuery.ORListQuery.FullNameList.Add(fullName);
            }
            return requestMsgSet;
        }
        private string[] getLineItem(int index)
        {
            string[] lineItem = new string[5];
            switch (index)
            {
                case 1:
                    lineItem[0] = (comboBox_Item1.Text == "" || comboBox_Item1.Text == null) ? "" : comboBox_Item1.Text;
                    lineItem[1] = (textBox_Desc1.Text == "" || textBox_Desc1.Text == null) ? "" : textBox_Desc1.Text;
                    lineItem[2] = (textBox_Qty1.Text == "" || textBox_Qty1.Text == null) ? "" : textBox_Qty1.Text;
                    lineItem[3] = (textBox_Price1.Text == "" || textBox_Price1.Text == null) ? "" : textBox_Price1.Text;
                    lineItem[4] = (textBox_Amount1.Text == "" || textBox_Amount1.Text == null) ? "" : textBox_Amount1.Text;
                    break;
                case 2:
                    lineItem[0] = (comboBox_Item2.Text == "" || comboBox_Item2.Text == null) ? "" : comboBox_Item2.Text;
                    lineItem[1] = (textBox_Desc2.Text == "" || textBox_Desc2.Text == null) ? "" : textBox_Desc2.Text;
                    lineItem[2] = (textBox_Qty2.Text == "" || textBox_Qty2.Text == null) ? "" : textBox_Qty2.Text;
                    lineItem[3] = (textBox_Price2.Text == "" || textBox_Price2.Text == null) ? "" : textBox_Price2.Text;
                    lineItem[4] = (textBox_Amount2.Text == "" || textBox_Amount2.Text == null) ? "" : textBox_Amount2.Text;
                    break;
                case 3:
                    lineItem[0] = (comboBox_Item3.Text == "" || comboBox_Item3.Text == null) ? "" : comboBox_Item3.Text;
                    lineItem[1] = (textBox_Desc3.Text == "" || textBox_Desc3.Text == null) ? "" : textBox_Desc3.Text;
                    lineItem[2] = (textBox_Qty3.Text == "" || textBox_Qty3.Text == null) ? "" : textBox_Qty3.Text;
                    lineItem[3] = (textBox_Price3.Text == "" || textBox_Price3.Text == null) ? "" : textBox_Price3.Text;
                    lineItem[4] = (textBox_Amount3.Text == "" || textBox_Amount3.Text == null) ? "" : textBox_Amount3.Text;
                    break;
                case 4:
                    lineItem[0] = (comboBox_Item4.Text == "" || comboBox_Item4.Text == null) ? "" : comboBox_Item4.Text;
                    lineItem[1] = (textBox_Desc4.Text == "" || textBox_Desc4.Text == null) ? "" : textBox_Desc4.Text;
                    lineItem[2] = (textBox_Qty4.Text == "" || textBox_Qty4.Text == null) ? "" : textBox_Qty4.Text;
                    lineItem[3] = (textBox_Price4.Text == "" || textBox_Price4.Text == null) ? "" : textBox_Price4.Text;
                    lineItem[4] = (textBox_Amount4.Text == "" || textBox_Amount4.Text == null) ? "" : textBox_Amount4.Text;
                    break;
                case 5:
                    lineItem[0] = (comboBox_Item5.Text == "" || comboBox_Item5.Text == null) ? "" : comboBox_Item5.Text;
                    lineItem[1] = (textBox_Desc5.Text == "" || textBox_Desc5.Text == null) ? "" : textBox_Desc5.Text;
                    lineItem[2] = (textBox_Qty5.Text == "" || textBox_Qty5.Text == null) ? "" : textBox_Qty5.Text;
                    lineItem[3] = (textBox_Price5.Text == "" || textBox_Price5.Text == null) ? "" : textBox_Price5.Text;
                    lineItem[4] = (textBox_Amount5.Text == "" || textBox_Amount5.Text == null) ? "" : textBox_Amount5.Text;
                    break;
            }
            return lineItem;
        }
        private IMsgSetRequest buildInvoiceAddRq()
        {
            if (!validateInput()) return null;
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IInvoiceAdd invAdd = requestMsgSet.AppendInvoiceAddRq();
            
            // CustomerRef -> FullName
            if (comboBox_Customer.Text != "")
            {
                invAdd.CustomerRef.FullName.SetValue(comboBox_Customer.Text);
            }

            // TxnDate 
            DateTime DT_TxnDate = System.DateTime.Today;
            if (dateTimePicker1.Text != "")
            {
                DT_TxnDate = Convert.ToDateTime(dateTimePicker1.Text);
                invAdd.TxnDate.SetValue(DT_TxnDate);
            }

            // RefNumber 
            if (textBox_RefNumber.Text != "")
            {
                invAdd.RefNumber.SetValue(textBox_RefNumber.Text);
            }

            // BillAddress
            if (label_BillTo.Text != "")
            {
                string[] BillAddress = label_BillTo.Text.Split('\n');
                invAdd.BillAddress.Addr1.SetValue(BillAddress[0]);
                invAdd.BillAddress.Addr2.SetValue(BillAddress[1]);
                invAdd.BillAddress.Addr3.SetValue(BillAddress[2]);
                invAdd.BillAddress.City.SetValue(BillAddress[3]);
                invAdd.BillAddress.State.SetValue(BillAddress[4]);
                invAdd.BillAddress.PostalCode.SetValue(BillAddress[5]);
            }

            // TermsRef -> FullName 
            bool termsAvailable = false;
            if (comboBox_Terms.Text != "")
            {
                termsAvailable = true;
                invAdd.TermsRef.FullName.SetValue(comboBox_Terms.Text);
            }

            // DueDate 
            if (termsAvailable)
            {
                DateTime DT_DueDate = System.DateTime.Today;
                double dueInDays = getDueInDays();
                DT_DueDate = DT_TxnDate.AddDays(dueInDays);
                invAdd.DueDate.SetValue(DT_DueDate);
            }

            // CustomerMsgRef -> FullName 
            if (comboBox_CustomerMessage.Text != "")
            {
                invAdd.CustomerMsgRef.FullName.SetValue(comboBox_CustomerMessage.Text);
            }

            // ExchangeRate 
            if (textBox_ExchangeRate.Text != "")
            {
                float exRate = (float)System.Convert.ToSingle(textBox_ExchangeRate.Text);
                invAdd.ExchangeRate.SetValue(exRate);
            }

            //Line Items
            for (int x = 1; x < 6; x++)
            {
                string[] lineItem = getLineItem(x);
                if (lineItem[0] == "") { 
                    break; 
                }
                IORInvoiceLineAdd invLineAdd = invAdd.ORInvoiceLineAddList.Append();
                if (lineItem[0] != "")
                {
                    invLineAdd.InvoiceLineAdd.ItemRef.FullName.SetValue(lineItem[0]);
                }
                if (lineItem[1] != "")
                {
                    invLineAdd.InvoiceLineAdd.Desc.SetValue(lineItem[1]);
                }
                if (lineItem[2] != "")
                {
                    invLineAdd.InvoiceLineAdd.Quantity.SetValue(Convert.ToDouble(lineItem[2]));
                }
                if (lineItem[3] != "")
                {
                    invLineAdd.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(lineItem[3]));
                }
                if (lineItem[4] != "")
                {
                    invLineAdd.InvoiceLineAdd.Amount.SetValue(Convert.ToDouble(lineItem[4]));
                }
            }
            return requestMsgSet;
        }


        // RESPONSE PARSING
        private string[] parsePreferencesQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            string[] retVal = new string[count];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            string statusMessage = response.StatusMessage;
            string statusSeverity = response.StatusSeverity;
            //MessageBox.Show("statusCode = " + statusCode + "\nstatusMessage = " + statusMessage + "\nstatusSeverity = " + statusSeverity);
            if (statusCode == 0)
            {
                IPreferencesRet prefRet = response.Detail as IPreferencesRet;
                if (prefRet.MultiCurrencyPreferences != null)
                {
                    retVal[0] = Convert.ToString(prefRet.MultiCurrencyPreferences.IsMultiCurrencyOn.GetValue());
                    retVal[1] = prefRet.MultiCurrencyPreferences.HomeCurrencyRef.FullName.GetValue();
                }
            }
            return retVal;
        }
        private int getCount(string request)
        {
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildDataCountQuery(request));
            int count = parseRsForCount(responseMsgSet);
            return count;
        }
        private int parseRsForCount(IMsgSetResponse responseMsgSet)
        {
            int ret = -1;
            try
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                ret = response.retCount;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error encountered: " + e.Message);
                ret = -1;
            }
            return ret;
        }
        private string[] parseCustomerQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            /*
             <?xml version="1.0" ?> 
             <QBXML>
             <QBXMLMsgsRs>
             <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                 <CustomerRet>
                     <FullName>Abercrombie, Kristy</FullName> 
                 </CustomerRet>
             </CustomerQueryRs>
             </QBXMLMsgsRs>
             </QBXML>    
            */
            string[] retVal = new string[count];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                ICustomerRetList custRetList = response.Detail as ICustomerRetList;
                for (int i = 0; i < count; i++)
                {
                    string fullName = null;
                    if(custRetList.GetAt(i).FullName!=null)
                    {
                        fullName=custRetList.GetAt(i).FullName.GetValue().ToString();
                        if (fullName != null)
                        {
                            retVal[i] = fullName;
                        }
                    }
                    IAddress billAddress = null;
                    if (custRetList.GetAt(i).BillAddress != null)
                    {
                        billAddress = custRetList.GetAt(i).BillAddress;
                        string addr1 = "", addr2 = "", addr3 = "", addr4 = "", addr5 = "";
                        string city = "", state = "", postalcode = "";
                        if (billAddress != null)
                        {
                            if(billAddress.Addr1!=null) addr1 = billAddress.Addr1.GetValue().ToString();
                            if(billAddress.Addr1 !=null) addr1 = billAddress.Addr1.GetValue().ToString();
                            if (billAddress.Addr2 != null) addr2 = billAddress.Addr2.GetValue().ToString();
                            if (billAddress.Addr3 != null) addr3 = billAddress.Addr3.GetValue().ToString();
                            if (billAddress.Addr4 != null) addr4 = billAddress.Addr4.GetValue().ToString();
                            if (billAddress.Addr5 != null) addr5 = billAddress.Addr5.GetValue().ToString();
                            if (billAddress.City != null) city = billAddress.City.GetValue().ToString();
                            if (billAddress.State != null) state = billAddress.State.GetValue().ToString();
                            if (billAddress.PostalCode != null) postalcode = billAddress.PostalCode.GetValue().ToString();

                            retVal[i] = addr1 + "\r\n" + addr2 + "\r\n"
                                + addr3 + "\r\n" 
                                + city + "\r\n" + state + "\r\n" + postalcode;
                        }
                    }
                    //IAddress shipAddress = null;
                    //if (custRetList.GetAt(i).ShipAddress != null)
                    //{
                    //    shipAddress = custRetList.GetAt(i).ShipAddress;
                    //    string addr1 = "", addr2 = "", addr3 = "", addr4 = "", addr5 = "";
                    //    string city = "", state = "", postalcode = "";
                    //    if (shipAddress != null)
                    //    {
                    //        if (shipAddress.Addr1 != null) addr1 = shipAddress.Addr1.GetValue().ToString();
                    //        if (shipAddress.Addr1 != null) addr1 = shipAddress.Addr1.GetValue().ToString();
                    //        if (shipAddress.Addr2 != null) addr2 = shipAddress.Addr2.GetValue().ToString();
                    //        if (shipAddress.Addr3 != null) addr3 = shipAddress.Addr3.GetValue().ToString();
                    //        if (shipAddress.Addr4 != null) addr4 = shipAddress.Addr4.GetValue().ToString();
                    //        if (shipAddress.Addr5 != null) addr5 = shipAddress.Addr5.GetValue().ToString();
                    //        if (shipAddress.City != null) city = shipAddress.City.GetValue().ToString();
                    //        if (shipAddress.State != null) state = shipAddress.State.GetValue().ToString();
                    //        if (shipAddress.PostalCode != null) postalcode = shipAddress.PostalCode.GetValue().ToString();

                    //        // RESUME HERE
                    //        retVal[i] = addr1 + "\r\n" + addr2 + "\r\n"
                    //            + addr3 + "\r\n"
                    //            + city + "\r\n" + state + "\r\n" + postalcode;
                    //    }
                    //}
                    string currencyRef=null;
                    if (custRetList.GetAt(i).CurrencyRef != null)
                    {
                        currencyRef = custRetList.GetAt(i).CurrencyRef.FullName.GetValue().ToString();
                        if (currencyRef != null)
                        {
                            retVal[i] = currencyRef;
                        }
                    }
                }
            }
            return retVal;
        }
        private string[,] parseItemQueryRs(IMsgSetResponse responseMsgSet, int countOfRows, int arraySize)
        {
            /*
              <?xml version="1.0" ?> 
            - <QBXML>
            - <QBXMLMsgsRs>
            - <ItemQueryRs requestID="2" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
            - <ItemServiceRet>
                <ListID>20000-933272655</ListID> 
                <TimeCreated>1999-07-29T11:24:15-08:00</TimeCreated> 
                <TimeModified>2007-12-15T11:32:53-08:00</TimeModified> 
                <EditSequence>1197747173</EditSequence> 
                <Name>Installation</Name> 
                <FullName>Installation</FullName> 
                <IsActive>true</IsActive> 
                <Sublevel>0</Sublevel> 
            - 	<SalesTaxCodeRef>
                    <ListID>20000-999022286</ListID> 
                    <FullName>Non</FullName> 
                </SalesTaxCodeRef>
            - 	<SalesOrPurchase>
                    <Desc>Installation labor</Desc> 
                    <Price>35.00</Price> 
            - 		<AccountRef>
                        <ListID>190000-933270541</ListID> 
                        <FullName>Construction Income:Labor Income</FullName> 
                    </AccountRef>
                </SalesOrPurchase>
              </ItemServiceRet>
              </ItemQueryRs>
              </QBXMLMsgsRs>
              </QBXML>
            */

            string[,] retVal = new string[countOfRows, arraySize];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                if (response.Detail == null)
                {
                    return null;
                }
                ENResponseType responseType = (ENResponseType)response.Type.GetValue();
                IORItemRetList OR=null;
                if (responseType == ENResponseType.rtItemQueryRs)
                {
                    OR = response.Detail as IORItemRetList;
                }
                else
                {
                    return null;
                }
                for (int i = 0; i < countOfRows; i++)
                {
                    if (OR.GetAt(i) == null) break;
                    if(OR.GetAt(i).ItemServiceRet!=null)
                    {
                        string fullName = null, desc = null;
                        double price = 0.0;

                        if (OR.GetAt(i).ItemServiceRet.FullName != null)
                        {
                            fullName = OR.GetAt(i).ItemServiceRet.FullName.GetValue();
                            populateRetVal(ref retVal, i, 0, "fullName", fullName);
                        }
                        if (OR.GetAt(i).ItemServiceRet.ORSalesPurchase != null)
                        {
                            if (OR.GetAt(i).ItemServiceRet.ORSalesPurchase.SalesOrPurchase != null)
                            {
                                //Get value of Desc
                                if (OR.GetAt(i).ItemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc != null)
                                {
                                    desc = (string)OR.GetAt(i).ItemServiceRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue();
                                    populateRetVal(ref retVal, i, 0, "desc", desc);
                                }
                                if (OR.GetAt(i).ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice != null)
                                {
                                    if (OR.GetAt(i).ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price != null)
                                    {
                                        //Get value of Price
                                        price = (double)OR.GetAt(i).ItemServiceRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue();
                                        populateRetVal(ref retVal, i, 1, "price", price.ToString());
                                    }
                                }
                            }
                        }
                    }
                    
                    if (OR.GetAt(i).ItemNonInventoryRet!= null)
                    {
                        string fullName = null, desc = null;
                        double price = 0.0;

                        if (OR.GetAt(i).ItemNonInventoryRet.FullName != null)
                        {
                            fullName = OR.GetAt(i).ItemNonInventoryRet.FullName.GetValue();
                            populateRetVal(ref retVal, i, 0, "fullName", fullName);
                        }
                        if (OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase != null)
                        {
                            if (OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase != null)
                            {
                                //Get value of Desc
                                if (OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.Desc != null)
                                {
                                    desc = (string)OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue();
                                    populateRetVal(ref retVal, i, 0, "desc", desc);
                                }
                                if (OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.ORPrice != null)
                                {
                                    if (OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price != null)
                                    {
                                        //Get value of Price
                                        price = (double)OR.GetAt(i).ItemNonInventoryRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue();
                                        populateRetVal(ref retVal, i, 1, "price", price.ToString());
                                    }
                                }
                            }
                        }
                    }
                    if (OR.GetAt(i).ItemOtherChargeRet != null)
                    {
                        string fullName = null, desc = null;
                        double price = 0.0;

                        if (OR.GetAt(i).ItemOtherChargeRet.FullName != null)
                        {
                            fullName = OR.GetAt(i).ItemOtherChargeRet.FullName.GetValue();
                            populateRetVal(ref retVal, i, 0, "fullName", fullName);
                        }
                        if (OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase != null)
                        {
                            if (OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase.SalesOrPurchase != null)
                            {
                                //Get value of Desc
                                if (OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.Desc != null)
                                {
                                    desc = (string)OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.Desc.GetValue();
                                    populateRetVal(ref retVal, i, 0, "desc", desc);
                                }
                                if (OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.ORPrice != null)
                                {
                                    if (OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price != null)
                                    {
                                        //Get value of Price
                                        price = (double)OR.GetAt(i).ItemOtherChargeRet.ORSalesPurchase.SalesOrPurchase.ORPrice.Price.GetValue();
                                        populateRetVal(ref retVal, i, 1, "price", price.ToString());
                                    }
                                }
                            }
                        }
                    } 
                    if (OR.GetAt(i).ItemInventoryRet!= null)
                    {
                        string fullName = null, desc = null;
                        double price = 0.0;

                        if (OR.GetAt(i).ItemInventoryRet.FullName != null)
                        {
                            fullName = OR.GetAt(i).ItemInventoryRet.FullName.GetValue();
                            populateRetVal(ref retVal, i, 0, "fullName", fullName);
                        }
                        if (OR.GetAt(i).ItemInventoryRet.SalesDesc != null)
                        {
                            //Get value of Desc
                            desc = (string)OR.GetAt(i).ItemInventoryRet.SalesDesc.GetValue();
                            populateRetVal(ref retVal, i, 0, "desc", desc);
                        }
                        if (OR.GetAt(i).ItemInventoryRet.SalesPrice!= null)
                        {
                            //Get value of Price
                            price = (double)OR.GetAt(i).ItemInventoryRet.SalesPrice.GetValue();
                            populateRetVal(ref retVal, i, 1, "price", price.ToString());
                        }
                    }
                    if (OR.GetAt(i).ItemInventoryAssemblyRet!= null)
                    {
                        string fullName = null, desc = null;
                        double price = 0.0;

                        if (OR.GetAt(i).ItemInventoryAssemblyRet.FullName != null)
                        {
                            fullName = OR.GetAt(i).ItemInventoryAssemblyRet.FullName.GetValue();
                            populateRetVal(ref retVal, i, 0, "fullName", fullName);
                        }
                        if (OR.GetAt(i).ItemInventoryAssemblyRet.SalesDesc != null)
                        {
                            //Get value of Desc
                            desc = (string)OR.GetAt(i).ItemInventoryAssemblyRet.SalesDesc.GetValue();
                            populateRetVal(ref retVal, i, 0, "desc", desc);
                        }
                        if (OR.GetAt(i).ItemInventoryAssemblyRet.SalesPrice != null)
                        {
                            //Get value of Price
                            price = (double)OR.GetAt(i).ItemInventoryAssemblyRet.SalesPrice.GetValue();
                            populateRetVal(ref retVal, i, 1, "price", price.ToString());
                        }
                    }
                }
            }
            return retVal;
        }
        private void populateRetVal(ref string[,] retVal, int row, int col, string fieldName, string fieldValue)
        {
            switch (fieldName)
            {
                case "fullName":
                case "desc":
                    if (fieldValue != null)
                    {
                        retVal[row, col] = fieldValue;
                    }
                    break;
                case "price":
                    double price = Convert.ToDouble(fieldValue);
                    if (price != 0.0)
                    {
                        if (exchangeRate != 1.0)
                        {
                            price = price / exchangeRate;
                        }
                        retVal[row, col] = price.ToString("N2");
                    }
                    break;
            }
        }
        private string[] parseTermsQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            /*
            <?xml version="1.0" ?> 
            <QBXML>
            <QBXMLMsgsRs>
            <TermsQueryRs requestID="3" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                <StandardTermsRet>
                    <Name>1% 10 Net 30</Name> 
                </StandardTermsRet>
            </TermsQueryRs>
            </QBXMLMsgsRs>
            </QBXML>            
            */
            string[] retVal = new string[count];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                IORTermsRetList termsRetList = response.Detail as IORTermsRetList;
                for (int i = 0; i < count; i++)
                {
                    string name=null;
                    if (termsRetList.GetAt(i).StandardTermsRet != null)
                    {
                        name = termsRetList.GetAt(i).StandardTermsRet.Name.GetValue();
                        if (name != null)
                        {
                            retVal[i] = name;
                        }
                    }
                    if (termsRetList.GetAt(i).DateDrivenTermsRet != null)
                    {
                        name = termsRetList.GetAt(i).DateDrivenTermsRet.Name.GetValue();
                        if (name != null)
                        {
                            retVal[i] = name;
                        }
                    }
                }
            }
            return retVal;
        }
        private string[] parseSalesTaxCodeQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            /*
            <?xml version="1.0" ?> 
            <QBXML>
            <QBXMLMsgsRs>
            <SalesTaxCodeQueryRs requestID="3" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                <SalesTaxCodeRet>
                    <FullName>Tax</FullName> 
                </SalesTaxCodeRet>
            </SalesTaxCodeQueryRs>
            </QBXMLMsgsRs>
            </QBXML>            
            */
            string[] retVal = new string[count];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                ISalesTaxCodeRetList salesTaxCodeRetList = response.Detail as ISalesTaxCodeRetList;
                for (int i = 0; i < count; i++)
                {
                    string name = null;
                    if (salesTaxCodeRetList.GetAt(i).Name != null)
                    {
                        name = salesTaxCodeRetList.GetAt(i).Name.GetValue();
                        if (name != null)
                        {
                            retVal[i] = name;
                        }
                    }
                }
            }
            return retVal;
        }
        private string[] parseCustomerMsgQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            string[] retVal = new string[count];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                ICustomerMsgRetList custMsgRetList = response.Detail as ICustomerMsgRetList;
                for (int i = 0; i < count; i++)
                {
                    string name = null;
                    if (custMsgRetList.GetAt(i).Name != null)
                    {
                        name = custMsgRetList.GetAt(i).Name.GetValue();
                        if (name != null)
                        {
                            retVal[i] = name;
                        }
                    }
                }
            }
            return retVal;
        }
        private string[] parseCurrencyQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            string[] retVal = new string[count];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                ICurrencyRetList currRetList = response.Detail as ICurrencyRetList;
                for (int i = 0; i < count; i++)
                {
                    string exRate = null;
                    if (currRetList.GetAt(i).ExchangeRate != null)
                    {
                        exRate = currRetList.GetAt(i).ExchangeRate.GetValue().ToString();
                        if (exRate != null)
                        {
                            retVal[i] = exRate;
                        }
                    }
                }
            }
            return retVal;
        }
        private string[] parseInvoiceAddRs(IMsgSetResponse responseMsgSet)
        {
            string[] retVal = new string[3];
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            retVal[0] = response.StatusCode.ToString();
            retVal[1] = response.StatusSeverity;
            retVal[2] = response.StatusMessage;
            return retVal;
        }



        // CONNECTION TO QB
        private void connectToQB()
        {
            sessionManager = SessionManager.getInstance();
            maxVersion = sessionManager.QBsdkMajorVersion;
        }
        private IMsgSetResponse processRequestFromQB(IMsgSetRequest requestSet)
        {
            try
            {
                //MessageBox.Show(requestSet.ToXMLString());
                IMsgSetResponse responseSet=sessionManager.doRequest(true, ref requestSet);
                //MessageBox.Show(responseSet.ToXMLString());
                return responseSet;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }
        private void disconnectFromQB()
        {
            if (sessionManager != null)
            {
                try
                {
                    sessionManager.endSession();
                    sessionManager.closeConnection();
                    sessionManager = null;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }


        //UI
        private void showStatus(string text)
        {
            labelLoadStatus.Visible = true;
            labelLoadStatus.Text = text;
            this.Refresh();
        }
        private void hideStatus()
        {
            labelLoadStatus.Visible = false;
        }
        private void fillComboBox(ComboBox cbox, string[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (values[i] != null) cbox.Items.Add(values[i]);
            }
        }
        private void comboBox_Customer_SelectedValueChanged(object sender, EventArgs e)
        {
            showStatus("Loading customer info...");
            label_BillTo.Text = getBillShipTo(comboBox_Customer.Text, "BillAddress").Trim();
            label_ShipTo.Text = getBillShipTo(comboBox_Customer.Text, "ShipAddress").Trim();
            workCurrency = getCurrencyCode(comboBox_Customer.Text).Trim();
            label1_Currency.Text = workCurrency;
            label3_Currency.Text = homeCurrency;
            label4_Currency.Text = homeCurrency;
            textBox_ExchangeRate.Text = getExchangeRate(workCurrency);
            exchangeRate = Convert.ToDouble(textBox_ExchangeRate.Text);
            hideStatus();
        }
        private void comboBox_Item1_SelectedValueChanged(object sender, EventArgs e)
        {
            getItem(1);
        }
        private void comboBox_Item2_SelectedValueChanged(object sender, EventArgs e)
        {
            getItem(2);
        }
        private void comboBox_Item3_SelectedValueChanged(object sender, EventArgs e)
        {
            getItem(3);
        }
        private void comboBox_Item4_SelectedValueChanged(object sender, EventArgs e)
        {
            getItem(4);
        }
        private void comboBox_Item5_SelectedValueChanged(object sender, EventArgs e)
        {
            getItem(5);
        }
        private void textBox_Qty1_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount1.Text = calculateAmount(textBox_Qty1.Text, textBox_Price1.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Price1_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount1.Text = calculateAmount(textBox_Qty1.Text, textBox_Price1.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Qty2_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount2.Text = calculateAmount(textBox_Qty2.Text, textBox_Price2.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Price2_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount2.Text = calculateAmount(textBox_Qty2.Text, textBox_Price2.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Qty3_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount3.Text = calculateAmount(textBox_Qty3.Text, textBox_Price3.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Price3_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount3.Text = calculateAmount(textBox_Qty3.Text, textBox_Price3.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Qty4_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount4.Text = calculateAmount(textBox_Qty4.Text, textBox_Price4.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Price4_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount4.Text = calculateAmount(textBox_Qty4.Text, textBox_Price4.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Qty5_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount5.Text = calculateAmount(textBox_Qty5.Text, textBox_Price5.Text);
            textBox_Total.Text = calculateTotal();
        }
        private void textBox_Price5_TextChanged(object sender, EventArgs e)
        {
            textBox_Amount5.Text = calculateAmount(textBox_Qty5.Text, textBox_Price5.Text);
            textBox_Total.Text = calculateTotal();
        }
        private string calculateAmount(string quantity, string price)
        {
            if (quantity == null || quantity == "") quantity = "0";
            if (price == null || price == "") price = "0";
            decimal qty = Convert.ToDecimal(quantity);
            decimal prc = Convert.ToDecimal(price);
            decimal amount = qty * prc;
            return amount.ToString();
        }
        private string calculateTotal()
        {
            decimal amount1 = (textBox_Amount1.Text == "" || textBox_Amount1 == null) ? 0 : Convert.ToDecimal(textBox_Amount1.Text);
            decimal amount2 = (textBox_Amount2.Text == "" || textBox_Amount2 == null) ? 0 : Convert.ToDecimal(textBox_Amount2.Text);
            decimal amount3 = (textBox_Amount3.Text == "" || textBox_Amount3 == null) ? 0 : Convert.ToDecimal(textBox_Amount3.Text);
            decimal amount4 = (textBox_Amount4.Text == "" || textBox_Amount4 == null) ? 0 : Convert.ToDecimal(textBox_Amount4.Text);
            decimal amount5 = (textBox_Amount5.Text == "" || textBox_Amount5 == null) ? 0 : Convert.ToDecimal(textBox_Amount5.Text);
            decimal sum = amount1 + amount2 + amount3 + amount4 + amount5;
            return sum.ToString();
        }
        private void calculateBalance()
        {
            decimal total = (textBox_Total.Text == "" || textBox_Total.Text == null) ? 0 : Convert.ToDecimal(textBox_Total.Text);
            decimal exrate = (textBox_ExchangeRate.Text == "" || textBox_ExchangeRate.Text == null) ? 0 : Convert.ToDecimal(textBox_ExchangeRate.Text);
            decimal balance = 0.0m;
            if (total > 0.0m && exrate > 0.0m) balance = total * exrate;
            textBox_BalanceDue.Text = balance.ToString("N2");
        }
        private void textBox_Total_TextChanged(object sender, EventArgs e)
        {
            calculateBalance();
        }
        private void textBox_ExchangeRate_TextChanged(object sender, EventArgs e)
        {
            calculateBalance();
        }
        private bool validateInput()
        {
            if (comboBox_Customer.Text == null || comboBox_Customer.Text == "") return false;
            if (comboBox_Item1.Text == null || comboBox_Item1.Text == "") return false;
            return true;
        }
        private double getDueInDays()
        {
            double dueInDays = 0;
            switch (comboBox_Terms.Text)
            {
                case "Due on receipt":
                    dueInDays = 0;
                    break;
                case "Net 15":
                    dueInDays = 15;
                    break;
                case "Net 30":
                    dueInDays = 30;
                    break;
                case "Net 60":
                    dueInDays = 60;
                    break;
                default:
                    dueInDays = 0;
                    break;
            }
            return dueInDays;
        }
        private void button_Save_Click(object sender, EventArgs e)
        {
            connectToQB();
            IMsgSetRequest requestMsgSet = buildInvoiceAddRq();
            if (requestMsgSet == null)
            {
                MessageBox.Show("One of the input is missing. Double-check your entries and then click Save again.", "Error saving invoice");
                return;
            }
            MessageBox.Show(requestMsgSet.ToXMLString());
            IMsgSetResponse responseMsgSet = processRequestFromQB(requestMsgSet);
            MessageBox.Show(responseMsgSet.ToXMLString());
            disconnectFromQB();
            string[] status = new string[3];
            if (responseMsgSet != null) status = parseInvoiceAddRs(responseMsgSet);
            string msg = "";

            if (responseMsgSet != null & status[0] == "0")
            {
                msg = "Invoice was added successfully!";
            }
            else
            {
                msg = "Could not add invoice.";
            }

            msg = msg + "\n\n";
            msg = msg + "Status Code = " + status[0] + "\n";
            msg = msg + "Status Severity = " + status[1] + "\n";
            msg = msg + "Status Message = " + status[2] + "\n";
            MessageBox.Show(msg);
        }

    }
}
