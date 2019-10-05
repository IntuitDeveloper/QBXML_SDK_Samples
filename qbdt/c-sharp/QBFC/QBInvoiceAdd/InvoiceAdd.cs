using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Interop.QBFC13;

namespace InvoiceAdd
{
	/// <summary>
	///---------------------------------------------------------------------------
	/// InvoiceAdd, InputItem : implementation files
	///
	/// Description:  
	/// This sample demonstrates how to populate some of the user input with 
	/// results from querying QuickBooks, allowing user to input data and add an 
	/// invoice to the QuickBooks running in your desktop (same machine). 
	/// This sample also implements QBFC's automatic error recovery feature.
	/// 
	/// Created On: 11/19/2004
	/// Copyright © 2004-2013 Intuit Inc. All rights reserved.
	/// Use is subject to the terms specified at:
	///     http://developer.intuit.com/legal/devsite_tos.html
	///
	/// Change Log: 
	/// 11/22/2004 = Customer:Job, Terms, Customer Message are now populated from QB
	/// 11/24/2004 = Invoice items are now populated from QB
	/// 11/30/2004 = Added automatic error recovery feature
	/// 12/02/2004 = Cleanup and Commenting
    /// 06/10/2010 - Updated to QBFC10
	/// 08/18/2011 - Updated to QBFC11
	/// 04/14/2012 - Updated to QBFC12
	/// 08/09/2013 - Updated to QBFC13
	///---------------------------------------------------------------------------
	/// </summary>
	public class frm1_InvoiceAdd : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lbl2_CustomerJob;
		private System.Windows.Forms.Label lbl3_Invoice;
		private System.Windows.Forms.Label lbl4_BillTo;
		private System.Windows.Forms.Label lbl5_BillTo_Addr1;
		private System.Windows.Forms.Label lbl6_BillTo_Addr2;
		private System.Windows.Forms.Label lbl7_BillTo_Addr3;
		private System.Windows.Forms.Label lbl8_BillTo_Addr4;
		private System.Windows.Forms.Label lbl9_BillTo_City;
		private System.Windows.Forms.Label lbl10_BillTo_State;
		private System.Windows.Forms.Label lbl11_BillTo_Postal;
		private System.Windows.Forms.Label lbl12_BillTo_Country;
		private System.Windows.Forms.TextBox txt2_BillTo_Addr1;
		private System.Windows.Forms.TextBox txt3_BillTo_Addr2;
		private System.Windows.Forms.TextBox txt4_BillTo_Addr3;
		private System.Windows.Forms.TextBox txt5_BillTo_Addr4;
		private System.Windows.Forms.TextBox txt6_BillTo_City;
		private System.Windows.Forms.TextBox txt7_BillTo_State;
		private System.Windows.Forms.TextBox txt8_BillTo_Postal;
		private System.Windows.Forms.TextBox txt9_BillTo_Country;
		private System.Windows.Forms.Label lbl13_InvoiceDate;
		private System.Windows.Forms.Label lbl14_InvoiceNo;
		private System.Windows.Forms.DateTimePicker dtTmPickr1_InvoiceDate;
		private System.Windows.Forms.TextBox txt10_InvoiceNo;
		private System.Windows.Forms.Label lbl15_PONumber;
		private System.Windows.Forms.Label lbl16_Terms;
		private System.Windows.Forms.Label lbl17_DueDate;
		private System.Windows.Forms.TextBox txt11_PONumber;
		private System.Windows.Forms.ComboBox cmboBx1_Terms;
		private System.Windows.Forms.DateTimePicker dtTmPickr2_DueDate;
		private System.Windows.Forms.PictureBox picBx_Intuit;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		// Custom added variables
		private TextBox[] RowItems;
		private System.Windows.Forms.ColumnHeader Item;
		private System.Windows.Forms.ColumnHeader Desc;
		private System.Windows.Forms.ColumnHeader Rate;
		private System.Windows.Forms.ColumnHeader Quantity;
		private System.Windows.Forms.ColumnHeader Amount;
		private System.Windows.Forms.ColumnHeader Tax;
		private System.Windows.Forms.ListView listView1_InvoiceItems;
		public System.Windows.Forms.TextBox txt17_Tax;
		private System.Windows.Forms.TextBox txt13_Desc;
		public  System.Windows.Forms.TextBox txt12_Item;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lbl20_Desc;
		private System.Windows.Forms.Label lbl21_Rate;
		private System.Windows.Forms.Label lbl22_Qty;
		private System.Windows.Forms.Label lbl23_Amount;
		private System.Windows.Forms.Label lbl24_Tax;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lbl26_Total;
		private System.Windows.Forms.TextBox txt20_Total;
		private System.Windows.Forms.Button btn1_Send;
		private System.Windows.Forms.Button btn2_Exit;
		private System.Windows.Forms.Label lbl27_CustomerMessage;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.ComboBox cmboBx3_CustomerJob;
		private System.Windows.Forms.ComboBox cmboBx4_CustomerMessage;
		private System.Windows.Forms.Button btn3_Reset;
    private TextBox txt16_Amount;
    private TextBox txt15_Qty;
    private TextBox txt14_Rate;
    public int CurrentRow=0;

		public frm1_InvoiceAdd()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			
			// Get CustomerJobs, Terms, and CustomerMessage
			populateCustomerJob();
			populateTerms();
			populateCustomerMessage();
			
			// Linking InvoiceItem column text boxes to a Textbox array
			RowItems = new TextBox[] {txt12_Item, txt13_Desc, txt14_Rate, txt15_Qty, txt16_Amount, txt17_Tax};
			InitializeRowEditing(CurrentRow);

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		// Line up all the edit boxes along the correct listview vertical borders
		void InitializeRowEditing(int rowNum)
		{
			int sum = 0;
			for (int i = 0; i < listView1_InvoiceItems.Columns.Count; i++)
			{
				int theWidth = listView1_InvoiceItems.Columns[i].Width;
				int theHeight = listView1_InvoiceItems.Items[0].Bounds.Height;
				RowItems[i].SetBounds(sum + listView1_InvoiceItems.Bounds.X, listView1_InvoiceItems.Bounds.Y + theHeight*rowNum, theWidth, theHeight);		
				sum += theWidth;
			}
			if (rowNum==7) 
			{
				listView1_InvoiceItems.Scrollable = true;
			}
		}



		// Dumps the row of edit boxes into the list view they are
		// positioned at.
		void DumpRowToListView(int nRow)
		{
			for (int i = 0; i < listView1_InvoiceItems.Columns.Count; i++)
			{
				if (listView1_InvoiceItems.Items[nRow].SubItems.Count - 1  < i)
				{
					listView1_InvoiceItems.Items[nRow].SubItems.Add(RowItems[i].Text);
				}
				else
				{
					listView1_InvoiceItems.Items[nRow].SubItems[i].Text = RowItems[i].Text;
				}
 			}
		}


		// Dumps the row of the listview into edit boxes
		// positioned at.
		void DumpListViewToRow(int nRow)
		{
			for (int i = 0; i < listView1_InvoiceItems.Columns.Count; i++)
			{
				if (listView1_InvoiceItems.Items[nRow].SubItems.Count - 1  < i)
				{
					RowItems[i].Text = "";
				}
				else
				{
					RowItems[i].Text = listView1_InvoiceItems.Items[nRow].SubItems[i].Text;
				}
			}
		}



		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
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
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm1_InvoiceAdd));
      System.Windows.Forms.ListViewItem listViewItem22 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem23 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem24 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem25 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem26 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem27 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem28 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem29 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem30 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem31 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem32 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem33 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem34 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem35 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem36 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem37 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem38 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem39 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem40 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem41 = new System.Windows.Forms.ListViewItem("");
      System.Windows.Forms.ListViewItem listViewItem42 = new System.Windows.Forms.ListViewItem("");
      this.lbl2_CustomerJob = new System.Windows.Forms.Label();
      this.lbl3_Invoice = new System.Windows.Forms.Label();
      this.lbl4_BillTo = new System.Windows.Forms.Label();
      this.lbl5_BillTo_Addr1 = new System.Windows.Forms.Label();
      this.lbl6_BillTo_Addr2 = new System.Windows.Forms.Label();
      this.lbl7_BillTo_Addr3 = new System.Windows.Forms.Label();
      this.lbl8_BillTo_Addr4 = new System.Windows.Forms.Label();
      this.lbl9_BillTo_City = new System.Windows.Forms.Label();
      this.lbl10_BillTo_State = new System.Windows.Forms.Label();
      this.lbl11_BillTo_Postal = new System.Windows.Forms.Label();
      this.lbl12_BillTo_Country = new System.Windows.Forms.Label();
      this.txt2_BillTo_Addr1 = new System.Windows.Forms.TextBox();
      this.txt3_BillTo_Addr2 = new System.Windows.Forms.TextBox();
      this.txt4_BillTo_Addr3 = new System.Windows.Forms.TextBox();
      this.txt5_BillTo_Addr4 = new System.Windows.Forms.TextBox();
      this.txt6_BillTo_City = new System.Windows.Forms.TextBox();
      this.txt7_BillTo_State = new System.Windows.Forms.TextBox();
      this.txt8_BillTo_Postal = new System.Windows.Forms.TextBox();
      this.txt9_BillTo_Country = new System.Windows.Forms.TextBox();
      this.lbl13_InvoiceDate = new System.Windows.Forms.Label();
      this.lbl14_InvoiceNo = new System.Windows.Forms.Label();
      this.dtTmPickr1_InvoiceDate = new System.Windows.Forms.DateTimePicker();
      this.txt10_InvoiceNo = new System.Windows.Forms.TextBox();
      this.lbl15_PONumber = new System.Windows.Forms.Label();
      this.lbl16_Terms = new System.Windows.Forms.Label();
      this.lbl17_DueDate = new System.Windows.Forms.Label();
      this.txt11_PONumber = new System.Windows.Forms.TextBox();
      this.cmboBx1_Terms = new System.Windows.Forms.ComboBox();
      this.dtTmPickr2_DueDate = new System.Windows.Forms.DateTimePicker();
      this.picBx_Intuit = new System.Windows.Forms.PictureBox();
      this.Item = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.Desc = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.Rate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.Quantity = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.Amount = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.Tax = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
      this.listView1_InvoiceItems = new System.Windows.Forms.ListView();
      this.txt17_Tax = new System.Windows.Forms.TextBox();
      this.txt13_Desc = new System.Windows.Forms.TextBox();
      this.txt12_Item = new System.Windows.Forms.TextBox();
      this.label1 = new System.Windows.Forms.Label();
      this.lbl20_Desc = new System.Windows.Forms.Label();
      this.lbl21_Rate = new System.Windows.Forms.Label();
      this.lbl22_Qty = new System.Windows.Forms.Label();
      this.lbl23_Amount = new System.Windows.Forms.Label();
      this.lbl24_Tax = new System.Windows.Forms.Label();
      this.label2 = new System.Windows.Forms.Label();
      this.label3 = new System.Windows.Forms.Label();
      this.label4 = new System.Windows.Forms.Label();
      this.label5 = new System.Windows.Forms.Label();
      this.label6 = new System.Windows.Forms.Label();
      this.label7 = new System.Windows.Forms.Label();
      this.lbl26_Total = new System.Windows.Forms.Label();
      this.txt20_Total = new System.Windows.Forms.TextBox();
      this.btn1_Send = new System.Windows.Forms.Button();
      this.btn2_Exit = new System.Windows.Forms.Button();
      this.lbl27_CustomerMessage = new System.Windows.Forms.Label();
      this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
      this.cmboBx3_CustomerJob = new System.Windows.Forms.ComboBox();
      this.cmboBx4_CustomerMessage = new System.Windows.Forms.ComboBox();
      this.btn3_Reset = new System.Windows.Forms.Button();
      this.txt16_Amount = new System.Windows.Forms.TextBox();
      this.txt15_Qty = new System.Windows.Forms.TextBox();
      this.txt14_Rate = new System.Windows.Forms.TextBox();
      ((System.ComponentModel.ISupportInitialize)(this.picBx_Intuit)).BeginInit();
      this.SuspendLayout();
      // 
      // lbl2_CustomerJob
      // 
      this.lbl2_CustomerJob.Location = new System.Drawing.Point(8, 72);
      this.lbl2_CustomerJob.Name = "lbl2_CustomerJob";
      this.lbl2_CustomerJob.Size = new System.Drawing.Size(80, 16);
      this.lbl2_CustomerJob.TabIndex = 1;
      this.lbl2_CustomerJob.Text = "Customer:Job";
      // 
      // lbl3_Invoice
      // 
      this.lbl3_Invoice.Font = new System.Drawing.Font("Tahoma", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.lbl3_Invoice.ForeColor = System.Drawing.Color.Gray;
      this.lbl3_Invoice.Location = new System.Drawing.Point(0, 8);
      this.lbl3_Invoice.Name = "lbl3_Invoice";
      this.lbl3_Invoice.Size = new System.Drawing.Size(152, 32);
      this.lbl3_Invoice.TabIndex = 3;
      this.lbl3_Invoice.Text = "Invoice";
      this.lbl3_Invoice.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
      // 
      // lbl4_BillTo
      // 
      this.lbl4_BillTo.Location = new System.Drawing.Point(8, 104);
      this.lbl4_BillTo.Name = "lbl4_BillTo";
      this.lbl4_BillTo.Size = new System.Drawing.Size(48, 16);
      this.lbl4_BillTo.TabIndex = 4;
      this.lbl4_BillTo.Text = "Bill To";
      // 
      // lbl5_BillTo_Addr1
      // 
      this.lbl5_BillTo_Addr1.Location = new System.Drawing.Point(8, 128);
      this.lbl5_BillTo_Addr1.Name = "lbl5_BillTo_Addr1";
      this.lbl5_BillTo_Addr1.Size = new System.Drawing.Size(40, 16);
      this.lbl5_BillTo_Addr1.TabIndex = 5;
      this.lbl5_BillTo_Addr1.Text = "Addr1";
      // 
      // lbl6_BillTo_Addr2
      // 
      this.lbl6_BillTo_Addr2.Location = new System.Drawing.Point(8, 152);
      this.lbl6_BillTo_Addr2.Name = "lbl6_BillTo_Addr2";
      this.lbl6_BillTo_Addr2.Size = new System.Drawing.Size(40, 16);
      this.lbl6_BillTo_Addr2.TabIndex = 6;
      this.lbl6_BillTo_Addr2.Text = "Addr2";
      // 
      // lbl7_BillTo_Addr3
      // 
      this.lbl7_BillTo_Addr3.Location = new System.Drawing.Point(8, 176);
      this.lbl7_BillTo_Addr3.Name = "lbl7_BillTo_Addr3";
      this.lbl7_BillTo_Addr3.Size = new System.Drawing.Size(40, 16);
      this.lbl7_BillTo_Addr3.TabIndex = 7;
      this.lbl7_BillTo_Addr3.Text = "Addr3";
      // 
      // lbl8_BillTo_Addr4
      // 
      this.lbl8_BillTo_Addr4.Location = new System.Drawing.Point(8, 200);
      this.lbl8_BillTo_Addr4.Name = "lbl8_BillTo_Addr4";
      this.lbl8_BillTo_Addr4.Size = new System.Drawing.Size(40, 16);
      this.lbl8_BillTo_Addr4.TabIndex = 8;
      this.lbl8_BillTo_Addr4.Text = "Addr4";
      // 
      // lbl9_BillTo_City
      // 
      this.lbl9_BillTo_City.Location = new System.Drawing.Point(8, 224);
      this.lbl9_BillTo_City.Name = "lbl9_BillTo_City";
      this.lbl9_BillTo_City.Size = new System.Drawing.Size(40, 16);
      this.lbl9_BillTo_City.TabIndex = 9;
      this.lbl9_BillTo_City.Text = "City";
      // 
      // lbl10_BillTo_State
      // 
      this.lbl10_BillTo_State.Location = new System.Drawing.Point(168, 224);
      this.lbl10_BillTo_State.Name = "lbl10_BillTo_State";
      this.lbl10_BillTo_State.Size = new System.Drawing.Size(40, 16);
      this.lbl10_BillTo_State.TabIndex = 10;
      this.lbl10_BillTo_State.Text = "State";
      // 
      // lbl11_BillTo_Postal
      // 
      this.lbl11_BillTo_Postal.Location = new System.Drawing.Point(8, 248);
      this.lbl11_BillTo_Postal.Name = "lbl11_BillTo_Postal";
      this.lbl11_BillTo_Postal.Size = new System.Drawing.Size(68, 16);
      this.lbl11_BillTo_Postal.TabIndex = 11;
      this.lbl11_BillTo_Postal.Text = "Postal Code";
      // 
      // lbl12_BillTo_Country
      // 
      this.lbl12_BillTo_Country.Location = new System.Drawing.Point(168, 248);
      this.lbl12_BillTo_Country.Name = "lbl12_BillTo_Country";
      this.lbl12_BillTo_Country.Size = new System.Drawing.Size(44, 16);
      this.lbl12_BillTo_Country.TabIndex = 12;
      this.lbl12_BillTo_Country.Text = "Country";
      // 
      // txt2_BillTo_Addr1
      // 
      this.txt2_BillTo_Addr1.Location = new System.Drawing.Point(48, 120);
      this.txt2_BillTo_Addr1.Name = "txt2_BillTo_Addr1";
      this.txt2_BillTo_Addr1.Size = new System.Drawing.Size(264, 20);
      this.txt2_BillTo_Addr1.TabIndex = 13;
      // 
      // txt3_BillTo_Addr2
      // 
      this.txt3_BillTo_Addr2.Location = new System.Drawing.Point(48, 144);
      this.txt3_BillTo_Addr2.Name = "txt3_BillTo_Addr2";
      this.txt3_BillTo_Addr2.Size = new System.Drawing.Size(264, 20);
      this.txt3_BillTo_Addr2.TabIndex = 14;
      // 
      // txt4_BillTo_Addr3
      // 
      this.txt4_BillTo_Addr3.Location = new System.Drawing.Point(48, 168);
      this.txt4_BillTo_Addr3.Name = "txt4_BillTo_Addr3";
      this.txt4_BillTo_Addr3.Size = new System.Drawing.Size(264, 20);
      this.txt4_BillTo_Addr3.TabIndex = 15;
      // 
      // txt5_BillTo_Addr4
      // 
      this.txt5_BillTo_Addr4.Location = new System.Drawing.Point(48, 192);
      this.txt5_BillTo_Addr4.Name = "txt5_BillTo_Addr4";
      this.txt5_BillTo_Addr4.Size = new System.Drawing.Size(264, 20);
      this.txt5_BillTo_Addr4.TabIndex = 16;
      // 
      // txt6_BillTo_City
      // 
      this.txt6_BillTo_City.Location = new System.Drawing.Point(48, 216);
      this.txt6_BillTo_City.Name = "txt6_BillTo_City";
      this.txt6_BillTo_City.Size = new System.Drawing.Size(112, 20);
      this.txt6_BillTo_City.TabIndex = 17;
      // 
      // txt7_BillTo_State
      // 
      this.txt7_BillTo_State.Location = new System.Drawing.Point(216, 216);
      this.txt7_BillTo_State.Name = "txt7_BillTo_State";
      this.txt7_BillTo_State.Size = new System.Drawing.Size(96, 20);
      this.txt7_BillTo_State.TabIndex = 18;
      // 
      // txt8_BillTo_Postal
      // 
      this.txt8_BillTo_Postal.Location = new System.Drawing.Point(80, 240);
      this.txt8_BillTo_Postal.Name = "txt8_BillTo_Postal";
      this.txt8_BillTo_Postal.Size = new System.Drawing.Size(80, 20);
      this.txt8_BillTo_Postal.TabIndex = 19;
      // 
      // txt9_BillTo_Country
      // 
      this.txt9_BillTo_Country.Location = new System.Drawing.Point(216, 240);
      this.txt9_BillTo_Country.Name = "txt9_BillTo_Country";
      this.txt9_BillTo_Country.Size = new System.Drawing.Size(96, 20);
      this.txt9_BillTo_Country.TabIndex = 20;
      // 
      // lbl13_InvoiceDate
      // 
      this.lbl13_InvoiceDate.Location = new System.Drawing.Point(320, 128);
      this.lbl13_InvoiceDate.Name = "lbl13_InvoiceDate";
      this.lbl13_InvoiceDate.Size = new System.Drawing.Size(72, 23);
      this.lbl13_InvoiceDate.TabIndex = 21;
      this.lbl13_InvoiceDate.Text = "Invoice Date";
      // 
      // lbl14_InvoiceNo
      // 
      this.lbl14_InvoiceNo.Location = new System.Drawing.Point(320, 168);
      this.lbl14_InvoiceNo.Name = "lbl14_InvoiceNo";
      this.lbl14_InvoiceNo.Size = new System.Drawing.Size(64, 16);
      this.lbl14_InvoiceNo.TabIndex = 22;
      this.lbl14_InvoiceNo.Text = "Invoice No.";
      // 
      // dtTmPickr1_InvoiceDate
      // 
      this.dtTmPickr1_InvoiceDate.Location = new System.Drawing.Point(392, 120);
      this.dtTmPickr1_InvoiceDate.Name = "dtTmPickr1_InvoiceDate";
      this.dtTmPickr1_InvoiceDate.Size = new System.Drawing.Size(256, 20);
      this.dtTmPickr1_InvoiceDate.TabIndex = 23;
      // 
      // txt10_InvoiceNo
      // 
      this.txt10_InvoiceNo.Location = new System.Drawing.Point(392, 160);
      this.txt10_InvoiceNo.Name = "txt10_InvoiceNo";
      this.txt10_InvoiceNo.Size = new System.Drawing.Size(88, 20);
      this.txt10_InvoiceNo.TabIndex = 24;
      this.txt10_InvoiceNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      // 
      // lbl15_PONumber
      // 
      this.lbl15_PONumber.Location = new System.Drawing.Point(488, 168);
      this.lbl15_PONumber.Name = "lbl15_PONumber";
      this.lbl15_PONumber.Size = new System.Drawing.Size(72, 23);
      this.lbl15_PONumber.TabIndex = 25;
      this.lbl15_PONumber.Text = "P.O. Number";
      // 
      // lbl16_Terms
      // 
      this.lbl16_Terms.Location = new System.Drawing.Point(320, 208);
      this.lbl16_Terms.Name = "lbl16_Terms";
      this.lbl16_Terms.Size = new System.Drawing.Size(40, 16);
      this.lbl16_Terms.TabIndex = 26;
      this.lbl16_Terms.Text = "Terms";
      // 
      // lbl17_DueDate
      // 
      this.lbl17_DueDate.Location = new System.Drawing.Point(320, 240);
      this.lbl17_DueDate.Name = "lbl17_DueDate";
      this.lbl17_DueDate.Size = new System.Drawing.Size(56, 23);
      this.lbl17_DueDate.TabIndex = 27;
      this.lbl17_DueDate.Text = "Due Date";
      // 
      // txt11_PONumber
      // 
      this.txt11_PONumber.Location = new System.Drawing.Point(560, 160);
      this.txt11_PONumber.Name = "txt11_PONumber";
      this.txt11_PONumber.Size = new System.Drawing.Size(88, 20);
      this.txt11_PONumber.TabIndex = 28;
      this.txt11_PONumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      // 
      // cmboBx1_Terms
      // 
      this.cmboBx1_Terms.Location = new System.Drawing.Point(392, 200);
      this.cmboBx1_Terms.Name = "cmboBx1_Terms";
      this.cmboBx1_Terms.Size = new System.Drawing.Size(256, 21);
      this.cmboBx1_Terms.TabIndex = 29;
      this.cmboBx1_Terms.Text = "Please select one from list";
      this.cmboBx1_Terms.TextChanged += new System.EventHandler(this.cmboBx1_Terms_Leave);
      this.cmboBx1_Terms.Leave += new System.EventHandler(this.cmboBx1_Terms_Leave);
      // 
      // dtTmPickr2_DueDate
      // 
      this.dtTmPickr2_DueDate.Location = new System.Drawing.Point(392, 240);
      this.dtTmPickr2_DueDate.Name = "dtTmPickr2_DueDate";
      this.dtTmPickr2_DueDate.Size = new System.Drawing.Size(256, 20);
      this.dtTmPickr2_DueDate.TabIndex = 30;
      // 
      // picBx_Intuit
      // 
      this.picBx_Intuit.Image = ((System.Drawing.Image)(resources.GetObject("picBx_Intuit.Image")));
      this.picBx_Intuit.Location = new System.Drawing.Point(552, 0);
      this.picBx_Intuit.Name = "picBx_Intuit";
      this.picBx_Intuit.Size = new System.Drawing.Size(104, 24);
      this.picBx_Intuit.TabIndex = 32;
      this.picBx_Intuit.TabStop = false;
      // 
      // Item
      // 
      this.Item.Text = "Item";
      this.Item.Width = 135;
      // 
      // Desc
      // 
      this.Desc.Text = "Desc";
      this.Desc.Width = 258;
      // 
      // Rate
      // 
      this.Rate.Text = "Rate";
      this.Rate.Width = 56;
      // 
      // Quantity
      // 
      this.Quantity.Text = "Quantity";
      this.Quantity.Width = 63;
      // 
      // Amount
      // 
      this.Amount.Text = "Amount";
      this.Amount.Width = 64;
      // 
      // Tax
      // 
      this.Tax.Text = "Tax";
      this.Tax.Width = 63;
      // 
      // listView1_InvoiceItems
      // 
      this.listView1_InvoiceItems.BorderStyle = System.Windows.Forms.BorderStyle.None;
      this.listView1_InvoiceItems.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.Item,
            this.Desc,
            this.Rate,
            this.Quantity,
            this.Amount,
            this.Tax});
      this.listView1_InvoiceItems.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.listView1_InvoiceItems.GridLines = true;
      this.listView1_InvoiceItems.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
      this.listView1_InvoiceItems.HideSelection = false;
      this.listView1_InvoiceItems.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem22,
            listViewItem23,
            listViewItem24,
            listViewItem25,
            listViewItem26,
            listViewItem27,
            listViewItem28,
            listViewItem29,
            listViewItem30,
            listViewItem31,
            listViewItem32,
            listViewItem33,
            listViewItem34,
            listViewItem35,
            listViewItem36,
            listViewItem37,
            listViewItem38,
            listViewItem39,
            listViewItem40,
            listViewItem41,
            listViewItem42});
      this.listView1_InvoiceItems.LabelEdit = true;
      this.listView1_InvoiceItems.Location = new System.Drawing.Point(8, 304);
      this.listView1_InvoiceItems.Name = "listView1_InvoiceItems";
      this.listView1_InvoiceItems.Scrollable = false;
      this.listView1_InvoiceItems.Size = new System.Drawing.Size(640, 192);
      this.listView1_InvoiceItems.TabIndex = 33;
      this.listView1_InvoiceItems.UseCompatibleStateImageBehavior = false;
      this.listView1_InvoiceItems.View = System.Windows.Forms.View.Details;
      // 
      // txt17_Tax
      // 
      this.txt17_Tax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.txt17_Tax.Location = new System.Drawing.Point(584, 320);
      this.txt17_Tax.Name = "txt17_Tax";
      this.txt17_Tax.Size = new System.Drawing.Size(64, 20);
      this.txt17_Tax.TabIndex = 45;
      this.txt17_Tax.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.txt17_Tax.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frm1_InvoiceAdd_KeyDown);
      this.txt17_Tax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frm1_InvoiceAdd_KeyPress);
      this.txt17_Tax.Leave += new System.EventHandler(this.txt17_Tax_Leave);
      // 
      // txt13_Desc
      // 
      this.txt13_Desc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.txt13_Desc.Location = new System.Drawing.Point(144, 320);
      this.txt13_Desc.Name = "txt13_Desc";
      this.txt13_Desc.Size = new System.Drawing.Size(256, 20);
      this.txt13_Desc.TabIndex = 41;
      // 
      // txt12_Item
      // 
      this.txt12_Item.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.txt12_Item.Cursor = System.Windows.Forms.Cursors.Arrow;
      this.txt12_Item.Location = new System.Drawing.Point(8, 320);
      this.txt12_Item.Name = "txt12_Item";
      this.txt12_Item.Size = new System.Drawing.Size(136, 20);
      this.txt12_Item.TabIndex = 40;
      this.txt12_Item.Text = "Double click to view list";
      this.txt12_Item.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.txt12_Item.DoubleClick += new System.EventHandler(this.txt12_Item_KeyDown);
      // 
      // label1
      // 
      this.label1.BackColor = System.Drawing.SystemColors.Control;
      this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.label1.Location = new System.Drawing.Point(8, 280);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(136, 23);
      this.label1.TabIndex = 46;
      this.label1.Text = "Item";
      this.label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
      // 
      // lbl20_Desc
      // 
      this.lbl20_Desc.BackColor = System.Drawing.SystemColors.Control;
      this.lbl20_Desc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.lbl20_Desc.Location = new System.Drawing.Point(144, 280);
      this.lbl20_Desc.Name = "lbl20_Desc";
      this.lbl20_Desc.Size = new System.Drawing.Size(256, 23);
      this.lbl20_Desc.TabIndex = 47;
      this.lbl20_Desc.Text = "Description";
      this.lbl20_Desc.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
      // 
      // lbl21_Rate
      // 
      this.lbl21_Rate.BackColor = System.Drawing.SystemColors.Control;
      this.lbl21_Rate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.lbl21_Rate.Location = new System.Drawing.Point(400, 280);
      this.lbl21_Rate.Name = "lbl21_Rate";
      this.lbl21_Rate.Size = new System.Drawing.Size(56, 23);
      this.lbl21_Rate.TabIndex = 48;
      this.lbl21_Rate.Text = "Rate";
      this.lbl21_Rate.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
      // 
      // lbl22_Qty
      // 
      this.lbl22_Qty.BackColor = System.Drawing.SystemColors.Control;
      this.lbl22_Qty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.lbl22_Qty.Location = new System.Drawing.Point(456, 280);
      this.lbl22_Qty.Name = "lbl22_Qty";
      this.lbl22_Qty.Size = new System.Drawing.Size(64, 23);
      this.lbl22_Qty.TabIndex = 49;
      this.lbl22_Qty.Text = "Quantity";
      this.lbl22_Qty.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
      // 
      // lbl23_Amount
      // 
      this.lbl23_Amount.BackColor = System.Drawing.SystemColors.Control;
      this.lbl23_Amount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.lbl23_Amount.Location = new System.Drawing.Point(520, 280);
      this.lbl23_Amount.Name = "lbl23_Amount";
      this.lbl23_Amount.Size = new System.Drawing.Size(64, 23);
      this.lbl23_Amount.TabIndex = 50;
      this.lbl23_Amount.Text = "Amount";
      this.lbl23_Amount.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
      // 
      // lbl24_Tax
      // 
      this.lbl24_Tax.BackColor = System.Drawing.SystemColors.Control;
      this.lbl24_Tax.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.lbl24_Tax.Location = new System.Drawing.Point(584, 280);
      this.lbl24_Tax.Name = "lbl24_Tax";
      this.lbl24_Tax.Size = new System.Drawing.Size(64, 23);
      this.lbl24_Tax.TabIndex = 51;
      this.lbl24_Tax.Text = "Tax (Y/N)";
      this.lbl24_Tax.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
      // 
      // label2
      // 
      this.label2.Location = new System.Drawing.Point(400, 288);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(48, 23);
      this.label2.TabIndex = 48;
      this.label2.Text = "Rate";
      // 
      // label3
      // 
      this.label3.Location = new System.Drawing.Point(456, 288);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(56, 23);
      this.label3.TabIndex = 49;
      this.label3.Text = "Quantity";
      // 
      // label4
      // 
      this.label4.Location = new System.Drawing.Point(520, 288);
      this.label4.Name = "label4";
      this.label4.Size = new System.Drawing.Size(56, 23);
      this.label4.TabIndex = 50;
      this.label4.Text = "Amount";
      // 
      // label5
      // 
      this.label5.Location = new System.Drawing.Point(584, 288);
      this.label5.Name = "label5";
      this.label5.Size = new System.Drawing.Size(56, 23);
      this.label5.TabIndex = 51;
      this.label5.Text = "Tax";
      // 
      // label6
      // 
      this.label6.Location = new System.Drawing.Point(8, 288);
      this.label6.Name = "label6";
      this.label6.Size = new System.Drawing.Size(100, 23);
      this.label6.TabIndex = 46;
      this.label6.Text = "Item";
      // 
      // label7
      // 
      this.label7.Location = new System.Drawing.Point(144, 288);
      this.label7.Name = "label7";
      this.label7.Size = new System.Drawing.Size(100, 23);
      this.label7.TabIndex = 47;
      this.label7.Text = "Description";
      // 
      // lbl26_Total
      // 
      this.lbl26_Total.Location = new System.Drawing.Point(440, 504);
      this.lbl26_Total.Name = "lbl26_Total";
      this.lbl26_Total.Size = new System.Drawing.Size(112, 16);
      this.lbl26_Total.TabIndex = 55;
      this.lbl26_Total.Text = "Total Due without tax";
      this.lbl26_Total.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
      // 
      // txt20_Total
      // 
      this.txt20_Total.Location = new System.Drawing.Point(560, 504);
      this.txt20_Total.Name = "txt20_Total";
      this.txt20_Total.Size = new System.Drawing.Size(88, 20);
      this.txt20_Total.TabIndex = 56;
      this.txt20_Total.Text = "0.00";
      this.txt20_Total.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      // 
      // btn1_Send
      // 
      this.btn1_Send.BackColor = System.Drawing.SystemColors.Control;
      this.btn1_Send.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.btn1_Send.Location = new System.Drawing.Point(488, 544);
      this.btn1_Send.Name = "btn1_Send";
      this.btn1_Send.Size = new System.Drawing.Size(80, 32);
      this.btn1_Send.TabIndex = 57;
      this.btn1_Send.Text = "Send";
      this.btn1_Send.UseVisualStyleBackColor = false;
      this.btn1_Send.Click += new System.EventHandler(this.btn1_Send_Click);
      // 
      // btn2_Exit
      // 
      this.btn2_Exit.BackColor = System.Drawing.SystemColors.Control;
      this.btn2_Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.btn2_Exit.Location = new System.Drawing.Point(576, 544);
      this.btn2_Exit.Name = "btn2_Exit";
      this.btn2_Exit.Size = new System.Drawing.Size(75, 32);
      this.btn2_Exit.TabIndex = 58;
      this.btn2_Exit.Text = "Exit";
      this.btn2_Exit.UseVisualStyleBackColor = false;
      this.btn2_Exit.Click += new System.EventHandler(this.btn2_Exit_Click);
      // 
      // lbl27_CustomerMessage
      // 
      this.lbl27_CustomerMessage.Location = new System.Drawing.Point(-8, 504);
      this.lbl27_CustomerMessage.Name = "lbl27_CustomerMessage";
      this.lbl27_CustomerMessage.Size = new System.Drawing.Size(96, 16);
      this.lbl27_CustomerMessage.TabIndex = 59;
      this.lbl27_CustomerMessage.Text = "Note to Customer";
      this.lbl27_CustomerMessage.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
      // 
      // cmboBx3_CustomerJob
      // 
      this.cmboBx3_CustomerJob.Location = new System.Drawing.Point(80, 64);
      this.cmboBx3_CustomerJob.Name = "cmboBx3_CustomerJob";
      this.cmboBx3_CustomerJob.Size = new System.Drawing.Size(232, 21);
      this.cmboBx3_CustomerJob.TabIndex = 61;
      this.cmboBx3_CustomerJob.Text = "Please select from list";
      this.cmboBx3_CustomerJob.TextChanged += new System.EventHandler(this.cmboBx3_CustomerJob_TextChanged);
      // 
      // cmboBx4_CustomerMessage
      // 
      this.cmboBx4_CustomerMessage.Location = new System.Drawing.Point(96, 504);
      this.cmboBx4_CustomerMessage.Name = "cmboBx4_CustomerMessage";
      this.cmboBx4_CustomerMessage.Size = new System.Drawing.Size(336, 21);
      this.cmboBx4_CustomerMessage.TabIndex = 62;
      // 
      // btn3_Reset
      // 
      this.btn3_Reset.BackColor = System.Drawing.SystemColors.Control;
      this.btn3_Reset.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.btn3_Reset.Location = new System.Drawing.Point(8, 544);
      this.btn3_Reset.Name = "btn3_Reset";
      this.btn3_Reset.Size = new System.Drawing.Size(80, 32);
      this.btn3_Reset.TabIndex = 63;
      this.btn3_Reset.Text = "New";
      this.btn3_Reset.UseVisualStyleBackColor = false;
      this.btn3_Reset.Click += new System.EventHandler(this.btn3_Reset_Click);
      // 
      // txt16_Amount
      // 
      this.txt16_Amount.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.txt16_Amount.Location = new System.Drawing.Point(520, 320);
      this.txt16_Amount.Name = "txt16_Amount";
      this.txt16_Amount.Size = new System.Drawing.Size(64, 20);
      this.txt16_Amount.TabIndex = 44;
      this.txt16_Amount.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.txt16_Amount.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.only_Numbers_KeyPress);
      // 
      // txt15_Qty
      // 
      this.txt15_Qty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.txt15_Qty.Location = new System.Drawing.Point(456, 320);
      this.txt15_Qty.Name = "txt15_Qty";
      this.txt15_Qty.Size = new System.Drawing.Size(64, 20);
      this.txt15_Qty.TabIndex = 43;
      this.txt15_Qty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.txt15_Qty.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.only_Numbers_KeyPress);
      this.txt15_Qty.Leave += new System.EventHandler(this.txt15_Qty_Leave);
      // 
      // txt14_Rate
      // 
      this.txt14_Rate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
      this.txt14_Rate.Location = new System.Drawing.Point(400, 320);
      this.txt14_Rate.Name = "txt14_Rate";
      this.txt14_Rate.Size = new System.Drawing.Size(56, 20);
      this.txt14_Rate.TabIndex = 42;
      this.txt14_Rate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      this.txt14_Rate.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.only_Numbers_KeyPress);
      // 
      // frm1_InvoiceAdd
      // 
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
      this.ClientSize = new System.Drawing.Size(656, 582);
      this.Controls.Add(this.btn3_Reset);
      this.Controls.Add(this.cmboBx4_CustomerMessage);
      this.Controls.Add(this.cmboBx3_CustomerJob);
      this.Controls.Add(this.lbl27_CustomerMessage);
      this.Controls.Add(this.btn2_Exit);
      this.Controls.Add(this.btn1_Send);
      this.Controls.Add(this.txt20_Total);
      this.Controls.Add(this.txt17_Tax);
      this.Controls.Add(this.txt16_Amount);
      this.Controls.Add(this.txt15_Qty);
      this.Controls.Add(this.txt14_Rate);
      this.Controls.Add(this.txt13_Desc);
      this.Controls.Add(this.txt12_Item);
      this.Controls.Add(this.txt11_PONumber);
      this.Controls.Add(this.txt10_InvoiceNo);
      this.Controls.Add(this.txt9_BillTo_Country);
      this.Controls.Add(this.txt8_BillTo_Postal);
      this.Controls.Add(this.txt7_BillTo_State);
      this.Controls.Add(this.txt6_BillTo_City);
      this.Controls.Add(this.txt5_BillTo_Addr4);
      this.Controls.Add(this.txt4_BillTo_Addr3);
      this.Controls.Add(this.txt3_BillTo_Addr2);
      this.Controls.Add(this.txt2_BillTo_Addr1);
      this.Controls.Add(this.lbl26_Total);
      this.Controls.Add(this.lbl24_Tax);
      this.Controls.Add(this.lbl23_Amount);
      this.Controls.Add(this.lbl22_Qty);
      this.Controls.Add(this.lbl21_Rate);
      this.Controls.Add(this.lbl20_Desc);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.listView1_InvoiceItems);
      this.Controls.Add(this.picBx_Intuit);
      this.Controls.Add(this.dtTmPickr2_DueDate);
      this.Controls.Add(this.cmboBx1_Terms);
      this.Controls.Add(this.lbl17_DueDate);
      this.Controls.Add(this.lbl16_Terms);
      this.Controls.Add(this.lbl15_PONumber);
      this.Controls.Add(this.dtTmPickr1_InvoiceDate);
      this.Controls.Add(this.lbl14_InvoiceNo);
      this.Controls.Add(this.lbl13_InvoiceDate);
      this.Controls.Add(this.lbl12_BillTo_Country);
      this.Controls.Add(this.lbl11_BillTo_Postal);
      this.Controls.Add(this.lbl10_BillTo_State);
      this.Controls.Add(this.lbl9_BillTo_City);
      this.Controls.Add(this.lbl8_BillTo_Addr4);
      this.Controls.Add(this.lbl7_BillTo_Addr3);
      this.Controls.Add(this.lbl6_BillTo_Addr2);
      this.Controls.Add(this.lbl5_BillTo_Addr1);
      this.Controls.Add(this.lbl4_BillTo);
      this.Controls.Add(this.lbl3_Invoice);
      this.Controls.Add(this.lbl2_CustomerJob);
      this.Controls.Add(this.label2);
      this.Controls.Add(this.label3);
      this.Controls.Add(this.label4);
      this.Controls.Add(this.label5);
      this.Controls.Add(this.label6);
      this.Controls.Add(this.label7);
      this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
      this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
      this.Name = "frm1_InvoiceAdd";
      this.Text = "Add an Invoice to QuickBooks";
      this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frm1_InvoiceAdd_KeyDown);
      this.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.frm1_InvoiceAdd_KeyPress);
      ((System.ComponentModel.ISupportInitialize)(this.picBx_Intuit)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new frm1_InvoiceAdd());
		}


		private void frm1_InvoiceAdd_KeyPress (object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			switch(e.KeyChar)
			{
				case 'a':
					break;
			}
		}

		private void txt17_Tax_Leave(object sender, System.EventArgs e)
		{
			HandleMoveDown();
		}

		void CalculateTotalForThisItem()
		{
			double sum=0;
			try
			{
				for (int i=0; i<listView1_InvoiceItems.Items.Count; i++)
				{
					if (listView1_InvoiceItems.Items[i].SubItems.Count == 6) // full row
					{
						try
						{
							double partSum=Convert.ToDouble(listView1_InvoiceItems.Items[i].SubItems[4].Text);
							sum += partSum;
							listView1_InvoiceItems.Items[i].SubItems[4].Text=partSum.ToString("#,###.00");
						}
						catch(FormatException ex)
						{
							sum += 0;
							String exceptionString = ex.Message.ToString();
						}
					}
				}
				txt20_Total.Text = sum.ToString("#,###.00");
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message.ToString());
			}
		}

		private void txt15_Qty_Leave(object sender, System.EventArgs e)
		{
			double rate=Convert.ToDouble(txt14_Rate.Text);
			double qty=Convert.ToDouble(txt15_Qty.Text);
			double amount=rate*qty;
			txt14_Rate.Text=rate.ToString("#,###.00");
			txt15_Qty.Text=qty.ToString("#,###.00");
			txt16_Amount.Text=amount.ToString("#,###.00");
		}


		private void cmboBx1_Terms_Leave(object sender, System.EventArgs e)
		{
			string invoiceDate=dtTmPickr1_InvoiceDate.Text;
			string terms=cmboBx1_Terms.Text;
			if (!invoiceDate.Equals("") || !terms.Equals(""))
			{
				System.DateTime invDate = Convert.ToDateTime(invoiceDate);
				System.DateTime dueDate = Convert.ToDateTime(dtTmPickr2_DueDate.Text);
				switch(terms.ToUpper().Trim())
				{
					case "DUE ON RECEIPT":
						dueDate=invDate.AddDays(0);
						break;
					case "NET 15":
						dueDate=invDate.AddDays(15);
						break;
					case "NET 30":
						dueDate=invDate.AddDays(30);
						break;
					case "NET 60":
						dueDate=invDate.AddDays(60);
						break;
				}
				dtTmPickr2_DueDate.Text=dueDate.ToString();
			}
		}


		void HandleMoveDown()
		{
			CalculateTotalForThisItem();
			DumpRowToListView(CurrentRow);
			CurrentRow++;
			if (CurrentRow > 20)
				CurrentRow=20;
			DumpListViewToRow(CurrentRow);
			InitializeRowEditing(CurrentRow);
			txt12_Item.Focus();
		}

		void HandleMoveUp()
		{
			DumpRowToListView(CurrentRow);
			CurrentRow--;
			if (CurrentRow < 0)
				CurrentRow = 0;
			DumpListViewToRow(CurrentRow);
			InitializeRowEditing(CurrentRow);					 
			txt12_Item.Focus();
		}

		private bool RowTextBoxesHaveFocus()
		{
			for (int i = 0; i < RowItems.Length; i++)
			{
				if (RowItems[i].Focused)
					return true;
			}
			return false;
		}

		private void frm1_InvoiceAdd_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			switch(e.KeyData.ToString().ToUpper())
			{
				case "TAB":
					if (txt17_Tax.Focused)
						goto case "Up";
					break;
				case "ENTER":
					if (RowTextBoxesHaveFocus())					
						goto case "Down";
					break;
				case "Down":
					HandleMoveDown();
					break;
				case "Up":
					HandleMoveUp();
					break;
			}
		}

		private void btn2_Exit_Click(object sender, System.EventArgs e)
		{
			Dispose();
		}

		private void btn1_Send_Click(object sender, System.EventArgs e)
		{
			QBFC_AddInvoice();
		}


 		// Code for handling different versions of QuickBooks
		private double QBFCLatestVersion(QBSessionManager SessionManager)
		{
			// Use oldest version to ensure that this application work with any QuickBooks (US)
			IMsgSetRequest msgset = SessionManager.CreateMsgSetRequest("US", 1, 0);
			msgset.AppendHostQueryRq();
			IMsgSetResponse QueryResponse = SessionManager.DoRequests(msgset);
			//MessageBox.Show("Host query = " + msgset.ToXMLString());
			//SaveXML(msgset.ToXMLString());

			
			// The response list contains only one response,
			// which corresponds to our single HostQuery request
			IResponse response = QueryResponse.ResponseList.GetAt(0);

			// Please refer to QBFC Developers Guide for details on why 
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
				}
			}
			return LastVers;
		}


		public IMsgSetRequest getLatestMsgSetRequest(QBSessionManager sessionManager)
		{
			// Find and adapt to supported version of QuickBooks
			double supportedVersion=QBFCLatestVersion(sessionManager);
			
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

			// Create the message set request object
			IMsgSetRequest requestMsgSet = sessionManager.CreateMsgSetRequest("US", qbXMLMajorVer, qbXMLMinorVer);
			return requestMsgSet;
		}


		public void populateItem(){
			InputItem Child = new InputItem();
			Child.Form1 = this; // Allow Child to access Form1 public members
			Child.Show();
		}
		
		public void populateCustomerJob(){

			// We want to know if we begun a session so we can end it if an
			// error happens
			bool booSessionBegun=false;
			
			// Create the session manager object using QBFC
			QBSessionManager sessionManager = new QBSessionManager();

			try
			{
				// Open the connection and begin a session to QuickBooks
				sessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
				//sessionManager.OpenConnection2("", "IDN InvoiceAdd C# sample",ENConnectionType.ctLocalQBDLaunchUI);
				sessionManager.BeginSession("", ENOpenMode.omDontCare);
				booSessionBegun=true;

				// Announcing QuickBooks version
				string QBVer=Convert.ToString(QBFCLatestVersion(sessionManager))+".0";
				MessageBox.Show("The qbXML version v" + QBVer + " is detected. Applicaton will set its compatibility accordingly." + 
					"\n\nThis sample uses QBFC for all of its communication to QuickBooks." + 
					"\n\nClick OK to Continue","Note",System.Windows.Forms.MessageBoxButtons.OK);

				// Get the RequestMsgSet based on the correct QB Version
				IMsgSetRequest requestSet = getLatestMsgSetRequest(sessionManager);
			
				// Initialize the message set request object
				requestSet.Attributes.OnError = ENRqOnError.roeStop;

				// Add the request to the message set request object
				ICustomerQuery CustQ=requestSet.AppendCustomerQueryRq();

				// Optionally, you can put filter on it.
				CustQ.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(50);

				// Do the request and get the response message set object
				IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

				// Uncomment the following to view and save the request and response XML
				// string requestXML = requestSet.ToXMLString();
				// MessageBox.Show(requestXML);
				// SaveXML(requestXML);
				// string responseXML = responseSet.ToXMLString();
				// MessageBox.Show(responseXML);
				// SaveXML(responseXML);

				IResponse response = responseSet.ResponseList.GetAt(0);
				// int statusCode = response.StatusCode;
				// string statusMessage = response.StatusMessage;
				// string statusSeverity = response.StatusSeverity;
				// MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);
			

				ICustomerRetList customerRetList = response.Detail as ICustomerRetList;
				if (!(customerRetList.Count==0))
				{
					for (int ndx=0; ndx<=(customerRetList.Count-1); ndx++)
					{
						ICustomerRet customerRet=customerRetList.GetAt(ndx);
						cmboBx3_CustomerJob.Items.Add(customerRet.FullName.GetValue());
					} // for
				} // if
	

				// Close the session and connection with QuickBooks
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


		public void populateBillTo()
		{

			// We want to know if we begun a session so we can end it if an
			// error happens
			bool booSessionBegun=false;

			if (!cmboBx3_CustomerJob.Text.Trim().StartsWith("Please select from list"))
			{
				// Create the session manager object using QBFC
				QBSessionManager sessionManager = new QBSessionManager();

				// Announcing QuickBooks version
				// string QBVer=Convert.ToString(QBFCLatestVersion(sessionManager))+".0";
				// MessageBox.Show("QuickBooks version v" + QBVer + " is detected. Applicaton will set its compatibility accordingly." + 
				//	"\n\nThis sample uses QBFC for all of its communication to QuickBooks." + 
				//	"\n\nClick OK to Continue","Note",System.Windows.Forms.MessageBoxButtons.OK);
				

				try
				{
					// Open the connection and begin a session to QuickBooks
					sessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
					sessionManager.BeginSession("", ENOpenMode.omDontCare);
					booSessionBegun=true;

					// Get the RequestMsgSet based on the correct QB Version
					IMsgSetRequest requestSet = getLatestMsgSetRequest(sessionManager);
				
					// Initialize the message set request object
					requestSet.Attributes.OnError = ENRqOnError.roeStop;

					// Add the request to the message set request object
					ICustomerQuery CustQ=requestSet.AppendCustomerQueryRq();

					/* Following VB example is taken from QBFC/DevGuide/pg100
					Example 3
					This query uses a list object to look for the employee records of two specific employees:
					Dim empQuery As QBFCxLib.IEmployeeQuery
					empQuery.ORListQuery.FullNameList.Add "Almira Smith"
					empQuery.ORListQuery.FullNameList.Add "Sisko Jones"
					*/
					string custFN = cmboBx3_CustomerJob.Text;
					CustQ.ORCustomerListQuery.FullNameList.Add(custFN);
					
					// Optionally, you can put filter on it.
					// CustQ.ORCustomerListQuery.CustomerListFilter.MaxReturned.SetValue(50);


					// Do the request and get the response message set object
					IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

					// Uncomment the following to view and save the request and response XML
					// string requestXML = requestSet.ToXMLString();
					// MessageBox.Show(requestXML);
					// SaveXML(requestXML);
					// string responseXML = responseSet.ToXMLString();
					// MessageBox.Show(responseXML);
					// SaveXML(responseXML);

					IResponse response = responseSet.ResponseList.GetAt(0);
					// int statusCode = response.StatusCode;
					// string statusMessage = response.StatusMessage;
					// string statusSeverity = response.StatusSeverity;
					// MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);
				

					ICustomerRetList customerRetList = response.Detail as ICustomerRetList;
					if (!(customerRetList.Count==0))
					{
						ICustomerRet customerRet=customerRetList.GetAt(0);
						if (customerRet.BillAddress != null)
						{
							if (customerRet.BillAddress.Addr1 != null)
							{
								txt2_BillTo_Addr1.Text=customerRet.BillAddress.Addr1.GetValue();
							}
							if (customerRet.BillAddress.Addr2 != null)
							{
								txt3_BillTo_Addr2.Text=customerRet.BillAddress.Addr2.GetValue();
							}
							if (customerRet.BillAddress.Addr3 != null)
							{
								txt4_BillTo_Addr3.Text=customerRet.BillAddress.Addr3.GetValue();
							}
							if (customerRet.BillAddress.Addr4 != null)
							{
								txt5_BillTo_Addr4.Text=customerRet.BillAddress.Addr4.GetValue();
							}
							if (customerRet.BillAddress.City != null)
							{
								txt6_BillTo_City.Text=customerRet.BillAddress.City.GetValue();
							}
							if (customerRet.BillAddress.State != null)
							{
								txt7_BillTo_State.Text=customerRet.BillAddress.State.GetValue();
							}
							if (customerRet.BillAddress.PostalCode != null)
							{
								txt8_BillTo_Postal.Text=customerRet.BillAddress.PostalCode.GetValue();
							}
							if (customerRet.BillAddress.Country != null)
							{	
								string country = customerRet.BillAddress.Country.GetValue();
								txt9_BillTo_Country.Text=country;
							}

						} // if BillAddress is not null
					} // if customeRetList.count not equals zero
		

					// Close the session and connection with QuickBooks
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
			} // if cmboBx3_CustomerJob.Text does not start with "Please select from list"
		}

		
		public void populateTerms()
		{
			// Create the session manager object using QBFC
			QBSessionManager sessionManager = new QBSessionManager();

			// We want to know if we begun a session so we can end it if an
			// error happens
			bool booSessionBegun=false;

			try
			{
				// Open the connection and begin a session to QuickBooks
				sessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
				sessionManager.BeginSession("", ENOpenMode.omDontCare);
				booSessionBegun=true;

				// Get the RequestMsgSet based on the correct QB Version
				IMsgSetRequest requestSet = getLatestMsgSetRequest(sessionManager);
			
				// Initialize the message set request object
				requestSet.Attributes.OnError = ENRqOnError.roeStop;

				// Add the request to the message set request object
				ITermsQuery TermsQ=requestSet.AppendTermsQueryRq();
				
				// Optionally, you can put filter on it.
				TermsQ.ORListQuery.ListFilter.MaxReturned.SetValue(10);

				// Do the request and get the response message set object
				IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

				// Uncomment the following to view and save the request and response XML
				// string requestXML = requestSet.ToXMLString();
				// MessageBox.Show(requestXML);
				// SaveXML(requestXML);
				// string responseXML = responseSet.ToXMLString();
				// MessageBox.Show(responseXML);
				// SaveXML(responseXML);

				IResponse response = responseSet.ResponseList.GetAt(0);
				// int statusCode = response.StatusCode;
				// string statusMessage = response.StatusMessage;
				// string statusSeverity = response.StatusSeverity;
				// MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);
				IORTermsRetList orTermsRetList = response.Detail as IORTermsRetList;

				if (!(orTermsRetList.Count==0))
				{
					for(int ndx=0; ndx<=(orTermsRetList.Count-1); ndx++)
					{ 
						IORTermsRet orTermsRet=orTermsRetList.GetAt(ndx);

						// The ortype property returns an enum
						// of the elements that can be contained in the OR object
						switch(orTermsRet.ortype)
						{
							case ENORTermsRet.ortrDateDrivenTermsRet:
								{
									IDateDrivenTermsRet DateDrivenTermsRet = orTermsRet.DateDrivenTermsRet;
									cmboBx1_Terms.Items.Add(DateDrivenTermsRet.Name.GetValue());
								}
								break;
							case ENORTermsRet.ortrStandardTermsRet:
								{
									IStandardTermsRet StandardTermsRet = orTermsRet.StandardTermsRet;
									cmboBx1_Terms.Items.Add(StandardTermsRet.Name.GetValue());
								}
								break;
						}
					} // for loop
				} // if 

				// Close the session and connection with QuickBooks
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
		

		public void populateCustomerMessage()
		{

			// Create the session manager object using QBFC
			QBSessionManager sessionManager = new QBSessionManager();

			// We want to know if we begun a session so we can end it if an
			// error happens
			bool booSessionBegun=false;

			try
			{
				// Open the connection and begin a session to QuickBooks
				sessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
				sessionManager.BeginSession("", ENOpenMode.omDontCare);
				booSessionBegun=true;

				// Get the RequestMsgSet based on the correct QB Version
				IMsgSetRequest requestSet = getLatestMsgSetRequest(sessionManager);
			
				// Initialize the message set request object
				requestSet.Attributes.OnError = ENRqOnError.roeStop;

				// Add the request to the message set request object
				ICustomerMsgQuery CustMsgQ=requestSet.AppendCustomerMsgQueryRq();

				// Optionally, you can put filter on it.
				CustMsgQ.ORListQuery.ListFilter.MaxReturned.SetValue(50);

				// Do the request and get the response message set object
				IMsgSetResponse responseSet = sessionManager.DoRequests(requestSet);

				// Uncomment the following to view and save the request and response XML
				// string requestXML = requestSet.ToXMLString();
				// MessageBox.Show(requestXML);
				// SaveXML(requestXML);
				// string responseXML = responseSet.ToXMLString();
				// MessageBox.Show(responseXML);
				// SaveXML(responseXML);

				IResponse response = responseSet.ResponseList.GetAt(0);
				// int statusCode = response.StatusCode;
				// string statusMessage = response.StatusMessage;
				// string statusSeverity = response.StatusSeverity;
				// MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);
			
				ICustomerMsgRetList customerMsgRetList = response.Detail as ICustomerMsgRetList;
				if (!(customerMsgRetList.Count==0))
				{
					for (int ndx=0; ndx<=(customerMsgRetList.Count-1); ndx++)
					{
						ICustomerMsgRet customerMsgRet=customerMsgRetList.GetAt(ndx);
						cmboBx4_CustomerMessage.Items.Add(customerMsgRet.Name.GetValue());
					} // for
				} // if

				// Close the session and connection with QuickBooks
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


		public void QBFC_AddInvoice()
		{
			IMsgSetRequest	requestMsgSet;
			IMsgSetResponse responseMsgSet;
			// Create the session manager object using QBFC
			QBSessionManager sessionManager = new QBSessionManager();

			// We want to know if we begun a session so we can end it if an
			// error happens
			bool booSessionBegun=false;

			try
			{
				// Use SessionManager object to open a connection and begin a session 
				// with QuickBooks. At this time, you should add interop.QBFCxLib into 
				// your Project References
				sessionManager.OpenConnection("", "IDN InvoiceAdd C# sample");
				sessionManager.BeginSession("", ENOpenMode.omDontCare);
				booSessionBegun=true;

				// Get the RequestMsgSet based on the correct QB Version
				requestMsgSet = getLatestMsgSetRequest(sessionManager);
				// requestMsgSet = sessionManager.CreateMsgSetRequest("US", 4, 0);

				// Initialize the message set request object
				requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
			
				// ERROR RECOVERY: 
				// All steps are described in QBFC Developers Guide, on pg 41
				// under section titled "Automated Error Recovery"

				// (1) Set the error recovery ID using ErrorRecoveryID function
				//		Value must be in GUID format
				//	You could use c:\Program Files\Microsoft Visual Studio\Common\Tools\GuidGen.exe 
				//	to create a GUID for your unique ID
				string errecid = "{E74068B5-0D6D-454d-B0FD-BDDF93CE67C3}";
				sessionManager.ErrorRecoveryID.SetValue(errecid);

				// (2) Set EnableErrorRecovery to true to enable error recovery
				sessionManager.EnableErrorRecovery = true;

				// (3) Set SaveAllMsgSetRequestInfo to true so the entire contents of the MsgSetRequest
				//		will be saved to disk. If SaveAllMsgSetRequestInfo is false (default), only the 
				//		newMessageSetID will be saved. 
				sessionManager.SaveAllMsgSetRequestInfo = true;

				// (4) Use IsErrorRecoveryInfo to check whether an unprocessed response exists. 
				//		If IsErrorRecoveryInfo is true:
				if (sessionManager.IsErrorRecoveryInfo())
				{
					//string reqXML;
					//string resXML;
					IMsgSetRequest reqMsgSet=null;
					IMsgSetResponse resMsgSet=null;

					// a. Get the response status, using GetErrorRecoveryStatus
					resMsgSet = sessionManager.GetErrorRecoveryStatus();
					// resXML = resMsgSet.ToXMLString();
					// MessageBox.Show(resXML);
					
					if (resMsgSet.Attributes.MessageSetStatusCode.Equals("600"))
					{
						// This case may occur when a transaction has failed after QB processed 
						// the request but client app didn't get the response and started with 
						// another company file.
						MessageBox.Show("The oldMessageSetID does not match any stored IDs, and no newMessageSetID is provided.");
					}
					else if (resMsgSet.Attributes.MessageSetStatusCode.Equals("9001"))
					{
						MessageBox.Show("Invalid checksum. The newMessageSetID specified, matches the currently stored ID, but checksum fails.");
					}
					else if (resMsgSet.Attributes.MessageSetStatusCode.Equals("9002"))
					{
						// Response was not successfully stored or stored properly
						MessageBox.Show("No stored response was found.");
					}
						// 9003 = Not used
					else if (resMsgSet.Attributes.MessageSetStatusCode.Equals("9004"))
					{
						// MessageSetID is set with a string of size > 24 char
						MessageBox.Show("Invalid MessageSetID, greater than 24 character was given.");
					}
					else if (resMsgSet.Attributes.MessageSetStatusCode.Equals("9005"))
					{
						MessageBox.Show("Unable to store response.");
					}
					else
					{	
						IResponse res = resMsgSet.ResponseList.GetAt(0);
						int sCode = res.StatusCode;
						//string sMessage = res.StatusMessage;
						//string sSeverity = res.StatusSeverity;
						//MessageBox.Show("StatusCode = " + sCode + "\n" + "StatusMessage = " + sMessage + "\n" + "StatusSeverity = " + sSeverity);
					
						if (sCode == 0)
						{
							MessageBox.Show("Last request was processed and Invoice was added successfully!");
						}
						else if (sCode > 0)
						{
							MessageBox.Show("There was a warning but last request was processed successfully!");
						}
						else
						{
							MessageBox.Show("It seems that there was an error in processing last request"); 
							// b. Get the saved request, using GetSavedMsgSetRequest
							reqMsgSet = sessionManager.GetSavedMsgSetRequest();
							//reqXML = reqMsgSet.ToXMLString();
							//MessageBox.Show(reqXML);
							
							// c. Process the response, possibly using the saved request
							resMsgSet = sessionManager.DoRequests(reqMsgSet);
							IResponse resp = resMsgSet.ResponseList.GetAt(0);
							int statCode = resp.StatusCode;
							if (statCode == 0)
							{	
								string resStr = null;
								IInvoiceRet invRet = resp.Detail as IInvoiceRet;
								resStr = resStr + "Following invoice has been successfully submitted to QuickBooks:\n\n\n";
								if (invRet.TxnNumber != null)
									resStr = resStr + "Txn Number = " + Convert.ToString(invRet.TxnNumber.GetValue()) + "\n";
							} // if (statusCode == 0)
						} // else (sCode)
					} // else (MessageSetStatusCode)
					
					// d. Clear the response status, using ClearErrorRecovery
					sessionManager.ClearErrorRecovery();
					MessageBox.Show("Proceeding with current transaction.");
				}
			
				
				// Add the request to the message set request object
				IInvoiceAdd invoiceAdd = requestMsgSet.AppendInvoiceAddRq();

				
				// Set the IInvoiceAdd field values
				// Customer:Job
				string customer = cmboBx3_CustomerJob.Text;
				if (!customer.Equals(""))
				{	
					invoiceAdd.CustomerRef.FullName.SetValue(customer);
				}
				// Invoice Date
				string invoiceDate=dtTmPickr1_InvoiceDate.Text;
				if (!invoiceDate.Equals(""))
				{
					invoiceAdd.TxnDate.SetValue(Convert.ToDateTime(invoiceDate));
				}
				// Invoice Number
				string invoiceNumber=txt10_InvoiceNo.Text;
				if (!invoiceNumber.Equals(""))
				{
					invoiceAdd.RefNumber.SetValue(invoiceNumber);
				}
				// Bill Address
				string bAddr1 = txt2_BillTo_Addr1.Text;
				string bAddr2 = txt3_BillTo_Addr2.Text;
				string bAddr3 = txt4_BillTo_Addr3.Text;
				string bAddr4 = txt5_BillTo_Addr4.Text;
				string bCity  = txt6_BillTo_City.Text;
				string bState = txt7_BillTo_State.Text;
				string bPostal = txt8_BillTo_Postal.Text;
				string bCountry = txt9_BillTo_Country.Text;
				invoiceAdd.BillAddress.Addr1.SetValue(bAddr1);
				invoiceAdd.BillAddress.Addr2.SetValue(bAddr2);
				invoiceAdd.BillAddress.Addr3.SetValue(bAddr3);
				invoiceAdd.BillAddress.Addr4.SetValue(bAddr4);
				invoiceAdd.BillAddress.City.SetValue(bCity);
				invoiceAdd.BillAddress.State.SetValue(bState);
				invoiceAdd.BillAddress.PostalCode.SetValue(bPostal);
				invoiceAdd.BillAddress.Country.SetValue(bCountry);
				// P.O. Number
				string poNumber = txt11_PONumber.Text;
				if (!poNumber.Equals(""))
				{
					invoiceAdd.PONumber.SetValue(poNumber);
				}
				// Terms
				string terms = cmboBx1_Terms.Text;
				if (terms.IndexOf("Please select one from list")>=0){
					terms="";
				}
				if (!terms.Equals(""))
				{
					invoiceAdd.TermsRef.FullName.SetValue(terms);
				}
				// Due Date
				string dueDate=dtTmPickr2_DueDate.Text;
				if (!dueDate.Equals(""))
				{
					invoiceAdd.DueDate.SetValue(Convert.ToDateTime(dueDate));
				}
				// Customer Message
				string customerMsg=cmboBx4_CustomerMessage.Text;
				if (!customerMsg.Equals(""))
				{
					invoiceAdd.CustomerMsgRef.FullName.SetValue(customerMsg);
				}
				
				// Set the values for the invoice line
				for (int i=0; i<listView1_InvoiceItems.Items.Count; i++)
				{
					// Create the line item for the invoice
					int c = listView1_InvoiceItems.Items[i].SubItems.Count;
					if (c == 6) // full row
					{
						string item = listView1_InvoiceItems.Items[i].SubItems[0].Text; // txt12_Item.Text;
						string desc = listView1_InvoiceItems.Items[i].SubItems[1].Text; // txt13_Desc.Text;
						string rate = listView1_InvoiceItems.Items[i].SubItems[2].Text; // txt14_Rate.Text;
						string qty  = listView1_InvoiceItems.Items[i].SubItems[3].Text; // txt15_Qty.Text;
						string amount = listView1_InvoiceItems.Items[i].SubItems[4].Text; // txt16_Amount.Text;
						string taxable = listView1_InvoiceItems.Items[i].SubItems[5].Text; // txt17_Tax.Text;
						
						if (!item.Equals("") || !desc.Equals(""))
						{
							IInvoiceLineAdd invoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append().InvoiceLineAdd;
							invoiceLineAdd.ItemRef.FullName.SetValue(item);
							invoiceLineAdd.Desc.SetValue(desc);
							invoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(rate));
							invoiceLineAdd.Quantity.SetValue(Convert.ToDouble(qty));
							invoiceLineAdd.Amount.SetValue(Convert.ToDouble(amount));
							// Currently IsTaxable is not supported in QBD - QuickBooks Desktop Edition
							/*
							if (taxable.ToUpper().Equals("Y") || taxable.ToUpper().Equals("N"))
							{
								bool isTaxable = false;
								if (taxable.ToUpper().Equals("Y")) isTaxable=true;
									invoiceLineAdd.IsTaxable.SetValue(isTaxable);
							}
							*/
						}
					}
				} // for

				// Uncomment the following to view and save the request and response XML
				// string requestXML = requestMsgSet.ToXMLString();
				// MessageBox.Show(requestXML);
				// SaveXML(requestXML);

					// If all inputs are in, perform the request and obtain a response from QuickBooks
					if (isAllInputIn())
					{
						responseMsgSet = sessionManager.DoRequests(requestMsgSet);
		
						// Uncomment the following to view and save the request and response XML
						string requestXML = requestMsgSet.ToXMLString();
						MessageBox.Show(requestXML);
						//SaveXML(requestXML);
						// string responseXML = responseSet.ToXMLString();
						// MessageBox.Show(responseXML);
						// SaveXML(responseXML);

						IResponse response = responseMsgSet.ResponseList.GetAt(0);
						int statusCode = response.StatusCode;
						// string statusMessage = response.StatusMessage;
						// string statusSeverity = response.StatusSeverity;
						// MessageBox.Show("Status:\nCode = " + statusCode + "\nMessage = " + statusMessage + "\nSeverity = " + statusSeverity);
				
						if (statusCode == 0)
						{	
							string resString = null;
							IInvoiceRet invoiceRet = response.Detail as IInvoiceRet;
							resString = resString + "Following invoice has been successfully submitted to QuickBooks:\n\n\n";
							if (invoiceRet.TimeCreated != null)
								resString = resString + "Time Created = " + Convert.ToString(invoiceRet.TimeCreated.GetValue()) + "\n";
							if (invoiceRet.TxnNumber != null)
								resString = resString + "Txn Number = " + Convert.ToString(invoiceRet.TxnNumber.GetValue()) + "\n";
							if (invoiceRet.TxnDate != null)
								resString = resString + "Txn Date = " + Convert.ToString(invoiceRet.TxnDate.GetValue()) + "\n";
							if (invoiceRet.RefNumber != null)
								resString = resString + "Reference Number = " + invoiceRet.RefNumber.GetValue() + "\n";
							if (invoiceRet.CustomerRef.FullName != null)
								resString = resString + "Customer FullName = " + invoiceRet.CustomerRef.FullName.GetValue() + "\n";
							resString = resString + "\nBilling Address:" + "\n";
							if (invoiceRet.BillAddress.Addr1 != null)
								resString = resString + "Addr1 = " + invoiceRet.BillAddress.Addr1.GetValue() + "\n";
							if (invoiceRet.BillAddress.Addr2 != null)
								resString = resString + "Addr2 = " + invoiceRet.BillAddress.Addr2.GetValue() + "\n";
							if (invoiceRet.BillAddress.Addr3 != null)
								resString = resString + "Addr3 = " + invoiceRet.BillAddress.Addr3.GetValue() + "\n";
							if (invoiceRet.BillAddress.Addr4 != null)
								resString = resString + "Addr4 = " + invoiceRet.BillAddress.Addr4.GetValue() + "\n";
							if (invoiceRet.BillAddress.City != null)
								resString = resString + "City = " + invoiceRet.BillAddress.City.GetValue() + "\n";
							if (invoiceRet.BillAddress.State != null)
								resString = resString + "State = " + invoiceRet.BillAddress.State.GetValue() + "\n";
							if (invoiceRet.BillAddress.PostalCode != null)
								resString = resString + "Postal Code = " + invoiceRet.BillAddress.PostalCode.GetValue() + "\n";
							if (invoiceRet.BillAddress.Country != null)
								resString = resString + "Country = " + invoiceRet.BillAddress.Country.GetValue() + "\n";
							if (invoiceRet.PONumber != null)
								resString = resString + "\nPO Number = " + invoiceRet.PONumber.GetValue() + "\n";
							if (invoiceRet.TermsRef.FullName != null)
								resString = resString + "Terms = " + invoiceRet.TermsRef.FullName.GetValue() + "\n";
							if (invoiceRet.DueDate != null)
								resString = resString + "Due Date = " + Convert.ToString(invoiceRet.DueDate.GetValue()) + "\n";
							if (invoiceRet.SalesTaxTotal != null)
								resString = resString + "Sales Tax = " + Convert.ToString(invoiceRet.SalesTaxTotal.GetValue()) + "\n";
							resString = resString + "\nInvoice Line Items:" + "\n";
							IORInvoiceLineRetList orInvoiceLineRetList = invoiceRet.ORInvoiceLineRetList;
							string fullname="<empty>";
							string desc="<empty>";
							string rate="<empty>";
							string quantity="<empty>";
							string amount="<empty>";
							for (int i=0; i<=orInvoiceLineRetList.Count-1; i++)
							{
								if(invoiceRet.ORInvoiceLineRetList.GetAt(i).ortype== ENORInvoiceLineRet.orilrInvoiceLineRet)
								{
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.ItemRef.FullName != null)
										fullname=invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.ItemRef.FullName.GetValue();
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.Desc != null)
										desc=invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.Desc.GetValue();
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.ORRate.Rate != null)
										rate=Convert.ToString(invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.ORRate.Rate.GetValue());
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.Quantity !=null)
										quantity=Convert.ToString(invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.Quantity.GetValue());
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.Amount != null)
										amount=Convert.ToString(invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineRet.Amount.GetValue());
								}
								else
								{
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.ItemGroupRef.FullName != null)
										fullname=invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.ItemGroupRef.FullName.GetValue();
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.Desc !=null)
										desc=invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.Desc.GetValue();
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.InvoiceLineRetList.GetAt(i).ORRate.Rate != null)
										rate=Convert.ToString(invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.InvoiceLineRetList.GetAt(i).ORRate.Rate.GetValue());
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.Quantity != null)
										quantity=Convert.ToString(invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.Quantity.GetValue());
									if (invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.TotalAmount != null)
										amount=Convert.ToString(invoiceRet.ORInvoiceLineRetList.GetAt(i).InvoiceLineGroupRet.TotalAmount.GetValue());
								}
								resString = resString + "Fullname: " + fullname + "\n";
								resString = resString + "Description: " + desc + "\n";
								resString = resString + "Rate: " + rate + "\n";
								resString = resString + "Quantity: " + quantity + "\n";
								resString = resString + "Amount: " + amount + "\n\n";
							}
							MessageBox.Show(resString);
						} // if statusCode is zero
					} // if all input is in
					else
					{
						MessageBox.Show("One or more required input is missing.\n\nPlease check and make sure all of the input have been entered.");
					}
				// Close the session and connection with QuickBooks
				sessionManager.ClearErrorRecovery();
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

		} // method QBFC_AddInvoice

		void SaveXML(string xmlstring)
		{
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				StreamWriter sr = new StreamWriter(saveFileDialog1.FileName);
				sr.Write(xmlstring);
				sr.Close();
			}
		}

		private void txt12_Item_KeyDown (object sender, System.EventArgs e)
		{
			populateItem();
		}

		private void cmboBx3_CustomerJob_TextChanged(object sender, System.EventArgs e)
		{
			populateBillTo();
		}

    private void only_Numbers_KeyPress(object sender, KeyPressEventArgs e)
    {
      if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
          (e.KeyChar != '.'))
      {
        e.Handled = true;
      }

      // only allow one decimal point
      if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
      {
        e.Handled = true;
      }
    }

    private void btn3_Reset_Click(object sender, System.EventArgs e)
		{
			cmboBx3_CustomerJob.Text = "Please select from list";
			txt2_BillTo_Addr1.Text = "";
			txt3_BillTo_Addr2.Text = "";
			txt4_BillTo_Addr3.Text = "";
			txt5_BillTo_Addr4.Text = "";
			txt6_BillTo_City.Text = "";
			txt7_BillTo_State.Text = "";
			txt8_BillTo_Postal.Text = "";
			txt9_BillTo_Country.Text = "";
			txt10_InvoiceNo.Text = "";
			txt11_PONumber.Text = "";
			dtTmPickr2_DueDate.Text = dtTmPickr1_InvoiceDate.Text;
			cmboBx1_Terms.Text = "Please select one from list";
			cmboBx4_CustomerMessage.Text = "";
			txt20_Total.Text = "0.00";
			// Input item
			InitializeRowEditing(0);
			if (listView1_InvoiceItems.Items.Count > 0)
			{
				for (int i = 0; i < listView1_InvoiceItems.Items.Count; i++)
				{
					listView1_InvoiceItems.Items.Clear();
				}
			}
			txt12_Item.Text="Double click to view list";
		}

		private bool isAllInputIn()
		{
			if (cmboBx3_CustomerJob.Text.Equals("") ||
				txt2_BillTo_Addr1.Text.Equals("") ||
				txt3_BillTo_Addr2.Text.Equals("") ||
				txt6_BillTo_City.Text.Equals("") ||
				txt7_BillTo_State.Text.Equals("") ||
				txt8_BillTo_Postal.Text.Equals("") ||
				txt10_InvoiceNo.Text.Equals("") ||
				txt11_PONumber.Text.Equals("") ||
				dtTmPickr2_DueDate.Text.Equals("") ||
				dtTmPickr1_InvoiceDate.Text.Equals("") ||
				cmboBx1_Terms.Text.Equals("") ||
				txt20_Total.Text.Equals(""))
			{
				return false;
			}
			else
			{
				return true;
			}
		}
  } // class frm1_InvoiceAdd
}
