
/*-----------------------------------------------------------
 * CustomerAddForm : implementation file
 *
 * Description:  This sample demonstrates the simple use 
 *               QuickBooks qbXMLRP COM object
 *				 Also it shows how to create and parse qbXML 	
 *				 using .NET XML classes
 *
 * Created On: 8/15/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *      http://developer.intuit.com/legal/devsite_tos.html
 *
 *----------------------------------------------------------
 */ 


using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Xml;
using Interop.QBXMLRP2;

namespace CustomerAdd
{
	/// <summary>
	/// CustomerAddForm shows how to invoke QuickBooks qbXMLRP COM object
	/// It uses .NET to create qbXML request and parse qbXML response
	/// </summary>
	public class CustomerAddForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button Exit;
		private System.Windows.Forms.Button AddCustomer;
		private System.Windows.Forms.TextBox Phone;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox CustName;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CustomerAddForm()
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
			this.label3 = new System.Windows.Forms.Label();
			this.Exit = new System.Windows.Forms.Button();
			this.AddCustomer = new System.Windows.Forms.Button();
			this.Phone = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.CustName = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(16, 96);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(200, 32);
			this.label3.TabIndex = 13;
			this.label3.Text = "Note: You need to have QuickBooks with a company file opened.";
			// 
			// Exit
			// 
			this.Exit.Location = new System.Drawing.Point(288, 112);
			this.Exit.Name = "Exit";
			this.Exit.Size = new System.Drawing.Size(120, 24);
			this.Exit.TabIndex = 12;
			this.Exit.Text = "Exit";
			this.Exit.Click += new System.EventHandler(this.Exit_Click);
			// 
			// AddCustomer
			// 
			this.AddCustomer.Location = new System.Drawing.Point(288, 80);
			this.AddCustomer.Name = "AddCustomer";
			this.AddCustomer.Size = new System.Drawing.Size(120, 24);
			this.AddCustomer.TabIndex = 11;
			this.AddCustomer.Text = "Add Customer";
			this.AddCustomer.Click += new System.EventHandler(this.AddCustomer_Click);
			// 
			// Phone
			// 
			this.Phone.Location = new System.Drawing.Point(96, 40);
			this.Phone.Name = "Phone";
			this.Phone.Size = new System.Drawing.Size(176, 20);
			this.Phone.TabIndex = 10;
			this.Phone.Text = "";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 16);
			this.label2.TabIndex = 9;
			this.label2.Text = "Phone";
			// 
			// CustName
			// 
			this.CustName.Location = new System.Drawing.Point(96, 16);
			this.CustName.Name = "CustName";
			this.CustName.Size = new System.Drawing.Size(312, 20);
			this.CustName.TabIndex = 8;
			this.CustName.Text = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 7;
			this.label1.Text = "Name";
			// 
			// CustomerAddForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(448, 149);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.label3,
																		  this.Exit,
																		  this.AddCustomer,
																		  this.Phone,
																		  this.label2,
																		  this.CustName,
																		  this.label1});
			this.Name = "CustomerAddForm";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "CustomerAdd";
			this.ResumeLayout(false);

		}
		#endregion
		static void Main() 
		{
			Application.Run(new CustomerAddForm());
		}

		private void Exit_Click(object sender, System.EventArgs e)
		{
			this.Close();

		}

		private void AddCustomer_Click(object sender, System.EventArgs e)
		{
			//step1: verify that Name is not empty
			String name = CustName.Text.Trim() ;
			if( name.Length == 0 )
			{
				MessageBox.Show("Please enter a value for Name.", "Input Validation");
				return;
			} 

			//step2: create the qbXML request
			XmlDocument inputXMLDoc = new XmlDocument();
			inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0",null, null));
			inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"2.0\""));
			XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
			inputXMLDoc.AppendChild(qbXML);
			XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
			qbXML.AppendChild(qbXMLMsgsRq);
			qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
			XmlElement custAddRq = inputXMLDoc.CreateElement("CustomerAddRq");
			qbXMLMsgsRq.AppendChild(custAddRq);
			custAddRq.SetAttribute("requestID", "1");
			XmlElement custAdd = inputXMLDoc.CreateElement("CustomerAdd");
			custAddRq.AppendChild(custAdd);	
			custAdd.AppendChild(inputXMLDoc.CreateElement("Name")).InnerText=name;
			if( Phone.Text.Length  > 0  )
			{
				custAdd.AppendChild(inputXMLDoc.CreateElement("Phone")).InnerText=Phone.Text;
			}
	
			string input = inputXMLDoc.OuterXml;
			//step3: do the qbXMLRP request
			RequestProcessor2 rp = null; 
			string ticket = null;
			string response = null;
			try 
			{
				rp = new RequestProcessor2 ();
				rp.OpenConnection("", "IDN CustomerAdd C# sample" );
				ticket = rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare );
				response = rp.ProcessRequest(ticket, input);
					
			}
			catch( System.Runtime.InteropServices.COMException ex )
			{
				MessageBox.Show( "COM Error Description = " +  ex.Message, "COM error" );
				return;
			}
			finally
			{
				if( ticket != null )
				{
					rp.EndSession(ticket);
				}
				if( rp != null )
				{
					rp.CloseConnection();
				}
			};

			//step4: parse the XML response and show a message
			XmlDocument outputXMLDoc = new XmlDocument();
			outputXMLDoc.LoadXml(response);
			XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName ("CustomerAddRs");
			
			if( qbXMLMsgsRsNodeList.Count == 1 ) //it's always true, since we added a single Customer
			{
				System.Text.StringBuilder popupMessage = new System.Text.StringBuilder();

				XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
				//get the status Code, info and Severity
				string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value; 
				string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
				string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;
                popupMessage.AppendFormat( "statusCode = {0}, statusSeverity = {1}, statusMessage = {2}",
					retStatusCode, retStatusSeverity, retStatusMessage );

				//get the CustomerRet node for detailed info
				
				//a CustomerAddRs contains max one childNode for "CustomerRet"
				XmlNodeList custAddRsNodeList = qbXMLMsgsRsNodeList.Item(0).ChildNodes;
				if( custAddRsNodeList.Count == 1 && custAddRsNodeList.Item(0).Name.Equals("CustomerRet") )
				{
					XmlNodeList custRetNodeList = custAddRsNodeList.Item(0).ChildNodes ;
				
					foreach (XmlNode custRetNode in custRetNodeList )
					{
						if ( custRetNode.Name.Equals( "ListID" ) )
						{
							popupMessage.AppendFormat("\r\nCustomer ListID = {0}", custRetNode.InnerText  );
						}
						else if ( custRetNode.Name.Equals( "Name" ) )
						{
							popupMessage.AppendFormat("\r\nCustomer Name = {0}", custRetNode.InnerText );
						}
						else if  ( custRetNode.Name.Equals( "FullName" ) )
						{
							popupMessage.AppendFormat("\r\nCustomer FullName = {0}", custRetNode.InnerText ); 
						}
					}
				} // End of customerRet

				MessageBox.Show (popupMessage.ToString(), "QuickBooks response");
			} //End of customerAddRs

		}
	}
}
