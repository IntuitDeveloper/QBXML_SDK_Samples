using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Interop.QBXMLRP2;

namespace MCInvoiceAdd {
    public partial class Form1 : Form {
        string workCurrency = "";
        string homeCurrency = "";
        decimal exchangeRate = 1.0m;

        public Form1() {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e) {
            this.Show();
            showStatus("Checking Multi Currency support...");
            if (!multiCurrencyOn())
            {
                string msg="Multi currency feature is not enabled for this company file. " +
                    "Use Edit -> Preferences -> Multiple Currencies to enable MultiCurrency " +
                    "support and then restart this sample.";
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
        
        private string ticket;
        private RequestProcessor2 rp;
        private string maxVersion;
        private string companyFile="";
        private QBFileMode mode=QBFileMode.qbFileOpenDoNotCare;
        
        private static string appID = "IDN123";
        private static string appName = "MCInvoiceAdd";

        private bool multiCurrencyOn(){
            bool ret = false;
            connectToQB();
            string response = processRequestFromQB(buildPreferencesQueryRqXML(new string[] { "MultiCurrencyPreferences" }, null));
            string[] prefs = parsePreferencesQueryRs(response, 2);
            ret = Convert.ToBoolean(prefs[0]);
            homeCurrency = prefs[1];
            disconnectFromQB();
            return ret;
        }

        private void loadCustomers() {
            string request = "CustomerQueryRq";
            connectToQB();
            int count = getCount(request);
            string response = processRequestFromQB(buildCustomerQueryRqXML(new string[] { "FullName" }, null));
            string[] customerList = parseCustomerQueryRs(response, count);
            disconnectFromQB();
            fillComboBox(this.comboBox_Customer, customerList);
        }

        private void loadItems() {
            string request = "ItemQueryRq";
            connectToQB();
            int count = getCount(request);
            string response = processRequestFromQB(buildItemQueryRqXML(new string[] { "FullName" }, null));
            string[] itemList = parseItemQueryRs(response, count);
            disconnectFromQB();
            fillComboBox(comboBox_Item1, itemList);
            fillComboBox(comboBox_Item2, itemList);
            fillComboBox(comboBox_Item3, itemList);
            fillComboBox(comboBox_Item4, itemList);
            fillComboBox(comboBox_Item5, itemList);
        }

        private void loadTerms() {
            string request = "TermsQueryRq";
            connectToQB();
            int count = getCount(request);
            string response = processRequestFromQB(buildTermsQueryRqXML());
            string[] termsList = parseTermsQueryRs(response, count);
            disconnectFromQB();
            fillComboBox(this.comboBox_Terms, termsList);
        }

        private void loadSalesTaxCodes() {
            string request = "SalesTaxCodeQueryRq";
            connectToQB();
            int count = getCount(request);
            string response = processRequestFromQB(buildSalesTaxCodeQueryRqXML());
            string[] salesTaxCodeList = parseSalesTaxCodeQueryRs(response, count);
            disconnectFromQB();
            fillComboBox(this.comboBox_Tax1, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax2, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax3, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax4, salesTaxCodeList);
            fillComboBox(this.comboBox_Tax5, salesTaxCodeList);
        }

        private string getBillShipTo(string customerName, string billOrShip) {
            connectToQB();
            string response=processRequestFromQB(buildCustomerQueryRqXML(new string[] { billOrShip }, customerName));
            string[] billShipTo = parseCustomerQueryRs(response, 1);
            if (billShipTo[0] == null) billShipTo[0] = "";
            disconnectFromQB();
            return billShipTo[0];
        }

        private string getCurrencyCode(string customerName){
            connectToQB();
            string response = processRequestFromQB(buildCustomerQueryRqXML(new string[] { "CurrencyRef" }, customerName));
            string[] currencyCode = parseCustomerQueryRs(response, 1);
            disconnectFromQB();
            return currencyCode[0];
        }

        private string getExchangeRate(string currencyName){
            connectToQB();
            string response = processRequestFromQB(buildCurrencyQueryRqXML(currencyName));
            string[] exrate = parseCurrencyQueryRs(response, 1);
            disconnectFromQB();
            if (exrate[0] == null || exrate[0] == "") exrate[0] = "1.0";
            return exrate[0];
        }
             
        private string[] getItemInfo(string itemName) {
            connectToQB();
            string response = processRequestFromQB(buildItemQueryRqXML(new string[] { "SalesOrPurchase" }, itemName));
            string[] itemInfo = parseItemQueryRs(response, 2);
            disconnectFromQB();
            return itemInfo;
        }

        private void loadCustomerMsg() {
            string request = "CustomerMsgQueryRq";
            connectToQB();
            int count = getCount(request);
            string response = processRequestFromQB(buildCustomerMsgQueryRqXML(new string[] { "Name" }, null));
            string[] customerMsgList = parseCustomerMsgQueryRs(response, count);
            disconnectFromQB();
            fillComboBox(comboBox_CustomerMessage, customerMsgList);
        }



        // RESPONSE PARSING
        private string[] parseInvoiceAddRs(string xml){
            string[] retVal = new string[3];
            try {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                RsNodeList = Doc.GetElementsByTagName("InvoiceAddRs");
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode statusCode = rsAttributes.GetNamedItem("statusCode");
                retVal[0] = Convert.ToString(statusCode.Value);
                XmlNode statusSeverity = rsAttributes.GetNamedItem("statusSeverity");
                retVal[1] = Convert.ToString(statusSeverity.Value);
                XmlNode statusMessage = rsAttributes.GetNamedItem("statusMessage");
                retVal[2] = Convert.ToString(statusMessage.Value);
            }
            catch (Exception e) {
                MessageBox.Show("Error encountered when parsing Invoice info returned from QuickBooks: " + e.Message);
                retVal = null;
            }
            return retVal;
        }

        private string[] parseCurrencyQueryRs(string xml, int count)
        {
            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CurrencyQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CurrencyRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ExchangeRate":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }
        
        private string[] parseCustomerMsgQueryRs(string xml, int count) {
            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null) {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more) {
                switch (nav.LocalName) {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CustomerMsgQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CustomerMsgRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Name":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }
       
        private string[] parseSalesTaxCodeQueryRs(string xml, int count) {
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
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null) {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more) {
                switch (nav.LocalName) {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "SalesTaxCodeQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "SalesTaxCodeRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Name":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        private string[] parseTermsQueryRs(string xml, int count) {
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
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null) {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more) {
                switch (nav.LocalName) {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "TermsQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "StandardTermsRet":
                    case "DateDrivenTermsRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Name":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        private string[] parsePreferencesQueryRs(string xml, int count)
        {
            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "PreferencesQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "PreferencesRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "MultiCurrencyPreferences":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "HomeCurrencyRef":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "IsMultiCurrencyOn":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        //more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        private string[] parseItemQueryRs(string xml, int count) {
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

            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null) {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more) {
                switch (nav.LocalName) {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemServiceRet":
                    case "ItemNonInventoryRet":
                    case "ItemOtherChargeRet":
                    case "ItemInventoryRet":
                    case "ItemInventoryAssemblyRet":
                    case "ItemFixedAssetRet":
                    case "ItemSubtotalRet":
                    case "ItemDiscountRet":
                    case "ItemPaymentRet":
                    case "ItemSalesTaxRet":
                    case "ItemSalesTaxGroupRet":
                    case "ItemGroupRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "SalesOrPurchase":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Desc":
                    case "Price":
                        string val=nav.Value.Trim();
                        decimal price=0.0m;
                        if (IsDecimal(val))
                        {
                            price = Convert.ToDecimal(val);
                            if (exchangeRate != 1.0m) { price = price / exchangeRate; }
                            retVal[1] = price.ToString("N2"); 
                        }
                        else
                        {
                            retVal[0] = val;
                        }
                        
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        private string[] parseCustomerQueryRs(string xml, int count) {
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
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null) {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more) {
                switch (nav.LocalName) {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CustomerQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CustomerRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "BillAddress":
                    case "ShipAddress":
                    case "CurrencyRef":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Addr1":
                    case "Addr2":
                    case "Addr3":
                    case "Addr4":
                    case "Addr5":
                    case "City":
                    case "State":
                    case "PostalCode":
                        retVal[x] = retVal[x] + "\r\n" + nav.Value.Trim();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        private int getCount(string request) {
            string response = processRequestFromQB(buildDataCountQuery(request));
            int count = parseRsForCount(response, request);
            return count;
        }

        public virtual int parseRsForCount(string xml, string request) {
            int ret = -1;
            try {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                string tagname = request.Replace("Rq", "Rs");
                RsNodeList = Doc.GetElementsByTagName(tagname);
                System.Text.StringBuilder popupMessage = new System.Text.StringBuilder();
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode retCount = rsAttributes.GetNamedItem("retCount");
                ret = Convert.ToInt32(retCount.Value);
            }
            catch (Exception e) {
                MessageBox.Show("Error encountered: " + e.Message);
                ret = -1;
            }
            return ret;
        }

        
        // REQUEST BUILDING
        
        //TODO: Build InvoiceAdd request xml.
        private string buildInvoiceAddRqXML()
        {
            string requestXML="";

            //if (!validateInput()) return null;

            //GET ALL INPUT INTO XML
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement InvoiceAddRq = xmlDoc.CreateElement("InvoiceAddRq");
            qbXMLMsgsRq.AppendChild(InvoiceAddRq);
            XmlElement InvoiceAdd = xmlDoc.CreateElement("InvoiceAdd");
            InvoiceAddRq.AppendChild(InvoiceAdd);
            
            // CustomerRef -> FullName
            if (comboBox_Customer.Text != "")
            {
                XmlElement Element_CustomerRef = xmlDoc.CreateElement("CustomerRef");
                InvoiceAdd.AppendChild(Element_CustomerRef);
                XmlElement Element_CustomerRef_FullName = xmlDoc.CreateElement("FullName");
                Element_CustomerRef.AppendChild(Element_CustomerRef_FullName).InnerText = comboBox_Customer.Text;
            }

            // TxnDate 
            DateTime DT_TxnDate = System.DateTime.Today;
            if (dateTimePicker1.Text != "")
            {
                DT_TxnDate = Convert.ToDateTime(dateTimePicker1.Text);
                string TxnDate = getDateString(DT_TxnDate);
                XmlElement Element_TxnDate = xmlDoc.CreateElement("TxnDate");
                InvoiceAdd.AppendChild(Element_TxnDate).InnerText = TxnDate;
            }

            // RefNumber 
            if (textBox_RefNumber.Text != "")
            {
                XmlElement Element_RefNumber = xmlDoc.CreateElement("RefNumber");
                InvoiceAdd.AppendChild(Element_RefNumber).InnerText = textBox_RefNumber.Text;
            }

            // BillAddress
            if (label_BillTo.Text != "")
            {
                string[] BillAddress = label_BillTo.Text.Split('\n');
                XmlElement Element_BillAddress = xmlDoc.CreateElement("BillAddress");
                InvoiceAdd.AppendChild(Element_BillAddress);
                for (int i = 0; i < BillAddress.Length; i++)
                {
                    if (BillAddress[i] != "" || BillAddress[i] != null)
                    {
                        XmlElement Element_Addr = xmlDoc.CreateElement("Addr" + (i + 1));
                        Element_BillAddress.AppendChild(Element_Addr).InnerText = BillAddress[i];
                    }
                }
            }

            // TermsRef -> FullName 
            bool termsAvailable = false;
            if (comboBox_Terms.Text != "")
            {
                termsAvailable = true;
                XmlElement Element_TermsRef = xmlDoc.CreateElement("TermsRef");
                InvoiceAdd.AppendChild(Element_TermsRef);
                XmlElement Element_TermsRef_FullName = xmlDoc.CreateElement("FullName");
                Element_TermsRef.AppendChild(Element_TermsRef_FullName).InnerText = comboBox_Terms.Text;
            }

            // DueDate 
            if (termsAvailable)
            {
                DateTime DT_DueDate = System.DateTime.Today;
                double dueInDays = getDueInDays();
                DT_DueDate = DT_TxnDate.AddDays(dueInDays);
                string DueDate = getDateString(DT_DueDate);
                XmlElement Element_DueDate = xmlDoc.CreateElement("DueDate");
                InvoiceAdd.AppendChild(Element_DueDate).InnerText = DueDate;
            }            
            
            // CustomerMsgRef -> FullName 
            if (comboBox_CustomerMessage.Text != "")
            {
                XmlElement Element_CustomerMsgRef = xmlDoc.CreateElement("CustomerMsgRef");
                InvoiceAdd.AppendChild(Element_CustomerMsgRef);
                XmlElement Element_CustomerMsgRef_FullName = xmlDoc.CreateElement("FullName");
                Element_CustomerMsgRef.AppendChild(Element_CustomerMsgRef_FullName).InnerText = comboBox_CustomerMessage.Text;
            }
            
            // ExchangeRate 
            if (textBox_ExchangeRate.Text != "")
            {
                XmlElement Element_ExchangeRate = xmlDoc.CreateElement("ExchangeRate");
                InvoiceAdd.AppendChild(Element_ExchangeRate).InnerText = textBox_ExchangeRate.Text;
            }

            //Line Items
            XmlElement Element_InvoiceLineAdd;
            for (int x = 1; x < 6; x++)
            {
                Element_InvoiceLineAdd = xmlDoc.CreateElement("InvoiceLineAdd");
                InvoiceAdd.AppendChild(Element_InvoiceLineAdd);

                string[] lineItem = getLineItem(x);
                if (lineItem[0] != "")
                {
                    XmlElement Element_InvoiceLineAdd_ItemRef = xmlDoc.CreateElement("ItemRef");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ItemRef);
                    XmlElement Element_InvoiceLineAdd_ItemRef_FullName = xmlDoc.CreateElement("FullName");
                    Element_InvoiceLineAdd_ItemRef.AppendChild(Element_InvoiceLineAdd_ItemRef_FullName).InnerText = lineItem[0];
                }
                if (lineItem[1] != "")
                {
                    XmlElement Element_InvoiceLineAdd_Desc = xmlDoc.CreateElement("Desc");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Desc).InnerText = lineItem[1];
                }
                if (lineItem[2] != "")
                {
                    XmlElement Element_InvoiceLineAdd_Quantity = xmlDoc.CreateElement("Quantity");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Quantity).InnerText = lineItem[2];
                }
                if (lineItem[3] != "")
                {
                    XmlElement Element_InvoiceLineAdd_Rate = xmlDoc.CreateElement("Rate");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Rate).InnerText = lineItem[3];
                }
                if (lineItem[4] != "")
                {
                    XmlElement Element_InvoiceLineAdd_Amount = xmlDoc.CreateElement("Amount");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Amount).InnerText = lineItem[4];
                }
            }
            
            
            InvoiceAddRq.SetAttribute("requestID", "99");
            requestXML = xmlDoc.OuterXml;

            return requestXML;
        }

        private string[] getLineItem(int index)
        {
            string[] lineItem = new string[5];
            switch (index)
            {
                case 1:
                    lineItem[0] = (comboBox_Item1.Text==""      || comboBox_Item1.Text==null)   ? "" : comboBox_Item1.Text;
                    lineItem[1] = (textBox_Desc1.Text == ""     || textBox_Desc1.Text == null)  ? "" : textBox_Desc1.Text;
                    lineItem[2] = (textBox_Qty1.Text == ""      || textBox_Qty1.Text == null)   ? "" : textBox_Qty1.Text;
                    lineItem[3] = (textBox_Price1.Text == ""    || textBox_Price1.Text == null) ? "" : textBox_Price1.Text;
                    lineItem[4] = (textBox_Amount1.Text == ""   || textBox_Amount1.Text == null)? "" : textBox_Amount1.Text;
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

        private bool validateInput()
        {
            if (comboBox_Customer.Text == null || comboBox_Customer.Text == "") return false;
            if (comboBox_Item1.Text == null || comboBox_Item1.Text == "") return false;
            return true;
        }

        private double getDueInDays()
        {
            double dueInDays=0;
            switch (comboBox_Terms.Text)
            {
                case "Due on receipt":
                    dueInDays=0;
                    break;
                case "Net 15":
                    dueInDays=15;
                    break;
                case "Net 30":
                    dueInDays=30;
                    break;
                case "Net 60":
                    dueInDays=60;
                    break;
                default:
                    dueInDays=0;
                    break;
            }
            return dueInDays;
        }

        private string getDateString(DateTime dt)
        {
            string year = dt.Year.ToString();
            string month = dt.Month.ToString();
            if (month.Length < 2) month = "0" + month;
            string day = dt.Day.ToString();
            if (day.Length < 2) day = "0" + day;
            return year + "-" + month + "-" + day;            
        }

        private string buildCurrencyQueryRqXML(string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CurrencyQueryRq = xmlDoc.CreateElement("CurrencyQueryRq");
            qbXMLMsgsRq.AppendChild(CurrencyQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CurrencyQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            CurrencyQueryRq.SetAttribute("requestID", "6");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        private string buildCustomerMsgQueryRqXML(string[] includeRetElement, string fullName) {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CustomerMsgQueryRq = xmlDoc.CreateElement("CustomerMsgQueryRq");
            qbXMLMsgsRq.AppendChild(CustomerMsgQueryRq);
            if (fullName != null) {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CustomerMsgQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++) {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                CustomerMsgQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            CustomerMsgQueryRq.SetAttribute("requestID", "5");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        private string buildSalesTaxCodeQueryRqXML() {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement salesTaxCodeQueryRq = xmlDoc.CreateElement("SalesTaxCodeQueryRq");
            qbXMLMsgsRq.AppendChild(salesTaxCodeQueryRq);
            XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
            salesTaxCodeQueryRq.AppendChild(includeRet).InnerText = "Name";
            salesTaxCodeQueryRq.SetAttribute("requestID", "4");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        private string buildTermsQueryRqXML() {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement termsQueryRq = xmlDoc.CreateElement("TermsQueryRq");
            qbXMLMsgsRq.AppendChild(termsQueryRq);
            XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
            termsQueryRq.AppendChild(includeRet).InnerText = "Name";
            termsQueryRq.SetAttribute("requestID", "3");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        private string buildItemQueryRqXML(string[] includeRetElement, string fullName) {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement ItemQueryRq = xmlDoc.CreateElement("ItemQueryRq");
            qbXMLMsgsRq.AppendChild(ItemQueryRq);
            if (fullName != null) {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                ItemQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++) {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                ItemQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            ItemQueryRq.SetAttribute("requestID", "2");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        private string buildPreferencesQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement PrefsQueryRq = xmlDoc.CreateElement("PreferencesQueryRq");
            qbXMLMsgsRq.AppendChild(PrefsQueryRq);
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                PrefsQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            PrefsQueryRq.SetAttribute("requestID", "1");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        private string buildCustomerQueryRqXML(string[] includeRetElement, string fullName) {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CustomerQueryRq = xmlDoc.CreateElement("CustomerQueryRq");
            qbXMLMsgsRq.AppendChild(CustomerQueryRq);
            if (fullName != null) {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CustomerQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++) {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                CustomerQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            CustomerQueryRq.SetAttribute("requestID", "1");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        public virtual string buildDataCountQuery(string request) {
            string input = "";
            XmlDocument inputXMLDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(inputXMLDoc, maxVersion);
            XmlElement queryRq = inputXMLDoc.CreateElement(request);
            queryRq.SetAttribute("metaData", "MetaDataOnly");
            qbXMLMsgsRq.AppendChild(queryRq);
            input = inputXMLDoc.OuterXml;
            return input;
        }

        private XmlElement buildRqEnvelope(XmlDocument doc, string maxVer) {
            doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
            doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"" + maxVer + "\""));
            XmlElement qbXML = doc.CreateElement("QBXML");
            doc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = doc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            return qbXMLMsgsRq;
        }

        
        // CONNECTION TO QB
        private void connectToQB() {
            rp = new RequestProcessor2Class();
            rp.OpenConnection(appID, appName);
            ticket = rp.BeginSession(companyFile, mode);
            string[] versions = rp.get_QBXMLVersionsForSession(ticket);
            maxVersion = versions[versions.Length - 1];
        }
        
        private void disconnectFromQB() {
            if (ticket != null) {
                try {
                    rp.EndSession(ticket);
                    ticket = null;
                    rp.CloseConnection();
                }
                catch (Exception e) {
                    MessageBox.Show(e.Message);
                }
            }
        }

        private string processRequestFromQB(string request) {
            try {
                return rp.ProcessRequest(ticket, request);
            }
            catch (Exception e) {
                MessageBox.Show(e.Message);
                return null;
            }
        }

        // Utilities
        private static bool IsDecimal(string theValue)
        {
            bool returnVal = false;
            try
            {
                Convert.ToDouble(theValue, System.Globalization.CultureInfo.CurrentCulture);
                returnVal = true;
            }
            catch 
            {
                returnVal = false;
            }
            finally
            {
            }

            return returnVal;

        }



        
        // UI
        private void showStatus(string text) {
            labelLoadStatus.Visible = true;
            labelLoadStatus.Text = text;
            this.Refresh();
        }
        
        private void hideStatus() {
            labelLoadStatus.Visible = false;
        }

        private void fillComboBox(ComboBox cbox, string[] values) {
            for (int i = 0; i < values.Length; i++) {
                if(values[i]!=null) cbox.Items.Add(values[i]);
            }
        }

        private void comboBox_Customer_SelectedValueChanged(object sender, EventArgs e) {
            showStatus("Loading customer info...");
            label_BillTo.Text = getBillShipTo(comboBox_Customer.Text, "BillAddress").Trim();
            label_ShipTo.Text = getBillShipTo(comboBox_Customer.Text, "ShipAddress").Trim();
            workCurrency = getCurrencyCode(comboBox_Customer.Text).Trim();
            label1_Currency.Text = workCurrency;
            label3_Currency.Text = homeCurrency;
            label4_Currency.Text = homeCurrency;
            textBox_ExchangeRate.Text = getExchangeRate(workCurrency);
            exchangeRate = Convert.ToDecimal(textBox_ExchangeRate.Text);
            hideStatus();
        }
        
        private void comboBox_Item1_SelectedValueChanged(object sender, EventArgs e) {
            getItem(1);
        }

        private void getItem(int index) {
            string[] itemInfo;
            switch (index) {
                case 1:
                    itemInfo=getItemInfo(comboBox_Item1.Text);
                    textBox_Desc1.Text=itemInfo[0];
                    textBox_Price1.Text=itemInfo[1];
                    break;
                case 2:
                    itemInfo = getItemInfo(comboBox_Item2.Text);
                    textBox_Desc2.Text = itemInfo[0];
                    textBox_Price2.Text = itemInfo[1];
                    break;
                case 3:
                    itemInfo = getItemInfo(comboBox_Item3.Text);
                    textBox_Desc3.Text = itemInfo[0];
                    textBox_Price3.Text = itemInfo[1];
                    break;
                case 4:
                    itemInfo = getItemInfo(comboBox_Item4.Text);
                    textBox_Desc4.Text = itemInfo[0];
                    textBox_Price4.Text = itemInfo[1];
                    break;
                case 5:
                    itemInfo = getItemInfo(comboBox_Item5.Text);
                    textBox_Desc5.Text = itemInfo[0];
                    textBox_Price5.Text = itemInfo[1];
                    break;
            }
            return;            
        }

        private string calculateAmount(string quantity, string price) {
            if (quantity == null || quantity == "") quantity = "0";
            if (price == null || price == "") price = "0";
            decimal qty = Convert.ToDecimal(quantity);
            decimal prc = Convert.ToDecimal(price);
            decimal amount = qty * prc;
            return amount.ToString();
        }

        private string calculateTotal() {
            decimal amount1 = (textBox_Amount1.Text == "" || textBox_Amount1 == null) ? 0 : Convert.ToDecimal(textBox_Amount1.Text);
            decimal amount2 = (textBox_Amount2.Text == "" || textBox_Amount2 == null) ? 0 : Convert.ToDecimal(textBox_Amount2.Text);
            decimal amount3 = (textBox_Amount3.Text == "" || textBox_Amount3 == null) ? 0 : Convert.ToDecimal(textBox_Amount3.Text);
            decimal amount4 = (textBox_Amount4.Text == "" || textBox_Amount4 == null) ? 0 : Convert.ToDecimal(textBox_Amount4.Text);
            decimal amount5 = (textBox_Amount5.Text == "" || textBox_Amount5 == null) ? 0 : Convert.ToDecimal(textBox_Amount5.Text);
            decimal sum = amount1 + amount2 + amount3 + amount4 + amount5;
            return sum.ToString();
        }

        private void textBox_Qty1_TextChanged(object sender, EventArgs e) {
            textBox_Amount1.Text = calculateAmount(textBox_Qty1.Text, textBox_Price1.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Price1_TextChanged(object sender, EventArgs e) {
            textBox_Amount1.Text = calculateAmount(textBox_Qty1.Text, textBox_Price1.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Qty2_TextChanged(object sender, EventArgs e) {
            textBox_Amount2.Text = calculateAmount(textBox_Qty2.Text, textBox_Price2.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Price2_TextChanged(object sender, EventArgs e) {
            textBox_Amount2.Text = calculateAmount(textBox_Qty2.Text, textBox_Price2.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Qty3_TextChanged(object sender, EventArgs e) {
            textBox_Amount3.Text = calculateAmount(textBox_Qty3.Text, textBox_Price3.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Price3_TextChanged(object sender, EventArgs e) {
            textBox_Amount3.Text = calculateAmount(textBox_Qty3.Text, textBox_Price3.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Qty4_TextChanged(object sender, EventArgs e) {
            textBox_Amount4.Text = calculateAmount(textBox_Qty4.Text, textBox_Price4.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Price4_TextChanged(object sender, EventArgs e) {
            textBox_Amount4.Text = calculateAmount(textBox_Qty4.Text, textBox_Price4.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Qty5_TextChanged(object sender, EventArgs e) {
            textBox_Amount5.Text = calculateAmount(textBox_Qty5.Text, textBox_Price5.Text);
            textBox_Total.Text = calculateTotal();
        }

        private void textBox_Price5_TextChanged(object sender, EventArgs e) {
            textBox_Amount5.Text = calculateAmount(textBox_Qty5.Text, textBox_Price5.Text);
            textBox_Total.Text = calculateTotal();
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

        private void button_Save_Click(object sender, EventArgs e)
        {
            string requestXML=buildInvoiceAddRqXML();
            if (requestXML == null)
            {
                MessageBox.Show("One of the input is missing. Double-check your entries and then click Save again.", "Error saving invoice");
                return;
            }
            connectToQB();
            string response = processRequestFromQB(requestXML);
            disconnectFromQB();
            string[] status = new string[3];
            if(response!=null) status=parseInvoiceAddRs(response);
            string msg = "";

            if (response!=null & status[0] == "0")
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