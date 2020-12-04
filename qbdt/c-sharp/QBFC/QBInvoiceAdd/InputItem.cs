using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using Interop.QBFC13;

namespace InvoiceAdd
{
	/// <summary>
	/// Summary description for InputItem.
	/// </summary>
	public class InputItem : System.Windows.Forms.Form
	{
		InvoiceAdd.frm1_InvoiceAdd form1;
			
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cmboBx2_Item;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btn1_OK;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public string isTaxable=null;

		public InputItem()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(InputItem));
			this.label1 = new System.Windows.Forms.Label();
			this.cmboBx2_Item = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.btn1_OK = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(264, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Below are the available items in QuickBooks:";
			// 
			// cmboBx2_Item
			// 
			this.cmboBx2_Item.Location = new System.Drawing.Point(8, 24);
			this.cmboBx2_Item.Name = "cmboBx2_Item";
			this.cmboBx2_Item.Size = new System.Drawing.Size(264, 21);
			this.cmboBx2_Item.TabIndex = 62;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(112, 16);
			this.label2.TabIndex = 63;
			this.label2.Text = "Click OK when done.";
			// 
			// btn1_OK
			// 
			this.btn1_OK.Location = new System.Drawing.Point(208, 48);
			this.btn1_OK.Name = "btn1_OK";
			this.btn1_OK.Size = new System.Drawing.Size(64, 24);
			this.btn1_OK.TabIndex = 64;
			this.btn1_OK.Text = "OK";
			this.btn1_OK.Click += new System.EventHandler(this.btn1_OK_Click);
			// 
			// InputItem
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			this.ClientSize = new System.Drawing.Size(280, 76);
			this.Controls.Add(this.btn1_OK);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmboBx2_Item);
			this.Controls.Add(this.label1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "InputItem";
			this.Text = "Select an item";
			this.Load += new System.EventHandler(this.InputItem_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void InputItem_Load(object sender, System.EventArgs e)
		{
			// IY: Create the session manager object using QBFC
			QBSessionManager sessionManager = new QBSessionManager();

			// IY: We want to know if we begun a session so we can end it if an
			// error happens
			bool booSessionBegun=false;

			try
			{
				// IY: Get the RequestMsgSet based on the correct QB Version
				IMsgSetRequest requestSet = getLatestMsgSetRequest(sessionManager);
			
				// IY: Initialize the message set request object
				requestSet.Attributes.OnError = ENRqOnError.roeStop;

				// IY: Add the request to the message set request object
				IItemQuery ItemQ=requestSet.AppendItemQueryRq();
				
				// IY: Optionally, you can put filter on it.
				// ItemQ.ORListQuery.ListFilter.MaxReturned.SetValue(30);

				// IY: Open the connection and begin a session to QuickBooks
				sessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
				sessionManager.BeginSession("", ENOpenMode.omDontCare);
				booSessionBegun=true;

				// IY: Do the request and get the response message set object
				IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

				// Uncomment the following to view and save the request and response XML
				//string requestXML = requestSet.ToXMLString();
				//MessageBox.Show(requestXML);
				// SaveXML(requestXML);
				//string responseXML = responseSet.ToXMLString();
				//MessageBox.Show(responseXML);
				// SaveXML(responseXML);

				IResponse response = responseSet.ResponseList.GetAt(0);
				//int statusCode = response.StatusCode;
				//string statusMessage = response.StatusMessage;
				//string statusSeverity = response.StatusSeverity;
				//MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);
				IORItemRetList orItemRetList = response.Detail as IORItemRetList;

				if (!(orItemRetList.Count==0))
				{
					for(int ndx=0; ndx<=(orItemRetList.Count-1); ndx++)
					{ 
						IORItemRet orItemRet=orItemRetList.GetAt(ndx);

						// IY: The ortype property returns an enum
						// of the elements that can be contained in the OR object
						switch(orItemRet.ortype)
						{
							case ENORItemRet.orirItemServiceRet:
							{
								// orir prefix comes from OR + Item + Ret
								IItemServiceRet ItemServiceRet = orItemRet.ItemServiceRet;
								isTaxable=ItemServiceRet.SalesTaxCodeRef.FullName.GetValue();
								cmboBx2_Item.Items.Add(ItemServiceRet.FullName.GetValue() + ":" + isTaxable);
							}
								break;
							case ENORItemRet.orirItemInventoryRet:
							{
								IItemInventoryRet ItemInventoryRet = orItemRet.ItemInventoryRet;
								isTaxable=ItemInventoryRet.SalesTaxCodeRef.FullName.GetValue();
								cmboBx2_Item.Items.Add(ItemInventoryRet.FullName.GetValue() + ":" + isTaxable);
							}
								break;
							case ENORItemRet.orirItemNonInventoryRet:
							{
								IItemNonInventoryRet ItemNonInventoryRet = orItemRet.ItemNonInventoryRet;
								isTaxable=ItemNonInventoryRet.SalesTaxCodeRef.FullName.GetValue();
								cmboBx2_Item.Items.Add(ItemNonInventoryRet.FullName.GetValue() + ":" + isTaxable);
							}
								break;
						}
					} // for loop
				} // if 

				// IY: Close the session and connection with QuickBooks
				sessionManager.EndSession();
				booSessionBegun=false;
				sessionManager.CloseConnection();
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message.ToString() + "\nStack Trace: \n" + ex.StackTrace + "\nExiting the application");
				if (booSessionBegun) 
				{
					sessionManager.EndSession();
					sessionManager.CloseConnection();
				}
			}

		}

		
		// IY: CODE FOR HANDLING DIFFERENT VERSIONS
		private double QBFCLatestVersion(QBSessionManager SessionManager)
		{
			// IY: Use oldest version to ensure that we work with any QuickBooks (US)
			IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
			msgset.AppendHostQueryRq();
			// MessageBox.Show(msgset.ToXMLString());

			// IY: Use SessionManager object to open a connection and begin a session 
			// with QuickBooks. At this time, you should add interop.QBFCxLib into 
			// your Project References
			SessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
			SessionManager.BeginSession("", ENOpenMode.omDontCare);

			IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);
			// IY: The response list contains only one response,
			// which corresponds to our single HostQuery request
			IResponse response = QueryResponse.ResponseList.GetAt(0);
			// IY: Please refer to QBFC Developers Guide/pg for details on why 
			// "as" clause was used to link this derrived class to its base class
			IHostRet HostResponse = response.Detail as IHostRet; 
			IBSTRList supportedVersions = HostResponse.SupportedQBXMLVersionList as IBSTRList;

			int i;
			double vers;
			double LastVers = 0;
			string svers = null;
			
			for (i=0; i<=supportedVersions.Count - 1; i++)
			{
				svers = supportedVersions.GetAt(i);
				vers=Convert.ToDouble(svers);
				if (vers > LastVers)
				{
					LastVers = vers;
					//svers = supportedVersions.GetAt(i);
				}
			}

			// IY: Close the session and connection with QuickBooks
			SessionManager.EndSession();
			SessionManager.CloseConnection();
			return LastVers;
		}


		public IMsgSetRequest getLatestMsgSetRequest(QBSessionManager sessionManager)
		{
			// IY: Find and adapt to supported version of QuickBooks
			double supportedVersion=QBFCLatestVersion(sessionManager);
			// MessageBox.Show("supportedVersion = " + supportedVersion.ToString());
			
			short qbXMLMajorVer=0;
			short qbXMLMinorVer=0;
			if (supportedVersion >=6.0) {
				qbXMLMajorVer=6;
				qbXMLMinorVer=0;
			}
			else if (supportedVersion >=5.0) {
				qbXMLMajorVer=5;
				qbXMLMinorVer=0;
			}
			else if (supportedVersion >=4.0)
			{
				qbXMLMajorVer=4;
				qbXMLMinorVer=0;
			}
			else if (supportedVersion >=3.0)
			{
				qbXMLMajorVer=3;
				qbXMLMinorVer=0;
			}
			else if (supportedVersion >=2.0)
			{
				qbXMLMajorVer=2;
				qbXMLMinorVer=0;
			}
			else if (supportedVersion >=1.1)
			{
				qbXMLMajorVer=1;
				qbXMLMinorVer=1;
			}
			else 
			{
				qbXMLMajorVer=1;
				qbXMLMinorVer=0;
				MessageBox.Show("It seems that you are running QuickBooks 2002 Release 1. We strongly recommend that you use QuickBooks' online update feature to obtain the latest fixes and enhancements");
			}
			// MessageBox.Show("qbXMLMajorVer = " + qbXMLMajorVer);
			// MessageBox.Show("qbXMLMinorVer = " + qbXMLMinorVer);

			// IY: Create the message set request object
			IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
			return requestMsgSet;
		}

		public InvoiceAdd.frm1_InvoiceAdd Form1 
		{
			get { return form1; }
			set { form1 = value; }
		}


		private void btn1_OK_Click(object sender, System.EventArgs e)
		{	
			string capText = cmboBx2_Item.Text;
			//MessageBox.Show("capText = " + capText);
			int pos = capText.IndexOf(":",0);
			//MessageBox.Show("pos = " + pos);
			string itemChosen = capText.Substring(0,pos);
			//MessageBox.Show("itemChosen = " + itemChosen);
			string isTax = capText.Substring(pos+1);
			//MessageBox.Show("isTax = " + isTax);
			form1.txt12_Item.Text = itemChosen;
			if (isTax.ToUpper().Equals("TAX"))	isTax="Y";
			else isTax="N";
			form1.txt17_Tax.Text=isTax;
			Close();
			Dispose();
		}


	} // public class
} // namespace
