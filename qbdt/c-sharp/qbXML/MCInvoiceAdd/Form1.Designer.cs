namespace MCInvoiceAdd {
    partial class Form1 {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.labelLoadStatus = new System.Windows.Forms.Label();
            this.comboBox_Customer = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_RefNumber = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label_BillTo = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label_ShipTo = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.comboBox_Terms = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.comboBox_Item1 = new System.Windows.Forms.ComboBox();
            this.textBox_Desc1 = new System.Windows.Forms.TextBox();
            this.textBox_Qty1 = new System.Windows.Forms.TextBox();
            this.textBox_Price1 = new System.Windows.Forms.TextBox();
            this.textBox_Amount1 = new System.Windows.Forms.TextBox();
            this.textBox_Amount2 = new System.Windows.Forms.TextBox();
            this.textBox_Price2 = new System.Windows.Forms.TextBox();
            this.textBox_Qty2 = new System.Windows.Forms.TextBox();
            this.textBox_Desc2 = new System.Windows.Forms.TextBox();
            this.comboBox_Item2 = new System.Windows.Forms.ComboBox();
            this.textBox_Amount3 = new System.Windows.Forms.TextBox();
            this.textBox_Price3 = new System.Windows.Forms.TextBox();
            this.textBox_Qty3 = new System.Windows.Forms.TextBox();
            this.textBox_Desc3 = new System.Windows.Forms.TextBox();
            this.comboBox_Item3 = new System.Windows.Forms.ComboBox();
            this.textBox_Amount4 = new System.Windows.Forms.TextBox();
            this.textBox_Price4 = new System.Windows.Forms.TextBox();
            this.textBox_Qty4 = new System.Windows.Forms.TextBox();
            this.textBox_Desc4 = new System.Windows.Forms.TextBox();
            this.comboBox_Item4 = new System.Windows.Forms.ComboBox();
            this.textBox_Amount5 = new System.Windows.Forms.TextBox();
            this.textBox_Price5 = new System.Windows.Forms.TextBox();
            this.textBox_Qty5 = new System.Windows.Forms.TextBox();
            this.textBox_Desc5 = new System.Windows.Forms.TextBox();
            this.comboBox_Item5 = new System.Windows.Forms.ComboBox();
            this.comboBox_Tax1 = new System.Windows.Forms.ComboBox();
            this.comboBox_Tax2 = new System.Windows.Forms.ComboBox();
            this.comboBox_Tax3 = new System.Windows.Forms.ComboBox();
            this.comboBox_Tax4 = new System.Windows.Forms.ComboBox();
            this.comboBox_Tax5 = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox_CustomerMessage = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.textBox_Total = new System.Windows.Forms.TextBox();
            this.label1_Currency = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.textBox_ExchangeRate = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.textBox_BalanceDue = new System.Windows.Forms.TextBox();
            this.label3_Currency = new System.Windows.Forms.Label();
            this.button_Save = new System.Windows.Forms.Button();
            this.label4_Currency = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelLoadStatus
            // 
            this.labelLoadStatus.AutoSize = true;
            this.labelLoadStatus.BackColor = System.Drawing.SystemColors.Control;
            this.labelLoadStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLoadStatus.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.labelLoadStatus.Location = new System.Drawing.Point(9, 481);
            this.labelLoadStatus.Name = "labelLoadStatus";
            this.labelLoadStatus.Size = new System.Drawing.Size(125, 20);
            this.labelLoadStatus.TabIndex = 0;
            this.labelLoadStatus.Text = "labelLoadStatus";
            this.labelLoadStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.labelLoadStatus.Visible = false;
            // 
            // comboBox_Customer
            // 
            this.comboBox_Customer.FormattingEnabled = true;
            this.comboBox_Customer.Location = new System.Drawing.Point(16, 20);
            this.comboBox_Customer.Name = "comboBox_Customer";
            this.comboBox_Customer.Size = new System.Drawing.Size(434, 21);
            this.comboBox_Customer.TabIndex = 1;
            this.comboBox_Customer.SelectedValueChanged += new System.EventHandler(this.comboBox_Customer_SelectedValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Customer:Job";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 44);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(108, 31);
            this.label2.TabIndex = 3;
            this.label2.Text = "Invoice";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(475, 21);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(190, 20);
            this.dateTimePicker1.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(472, 5);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(30, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Date";
            // 
            // textBox_RefNumber
            // 
            this.textBox_RefNumber.Location = new System.Drawing.Point(533, 48);
            this.textBox_RefNumber.Name = "textBox_RefNumber";
            this.textBox_RefNumber.Size = new System.Drawing.Size(132, 20);
            this.textBox_RefNumber.TabIndex = 6;
            this.textBox_RefNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(475, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Invoice #";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(19, 77);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(39, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Bill To:";
            // 
            // label_BillTo
            // 
            this.label_BillTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label_BillTo.Location = new System.Drawing.Point(17, 98);
            this.label_BillTo.Name = "label_BillTo";
            this.label_BillTo.Size = new System.Drawing.Size(190, 80);
            this.label_BillTo.TabIndex = 9;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(478, 77);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(47, 13);
            this.label7.TabIndex = 10;
            this.label7.Text = "Ship To:";
            // 
            // label_ShipTo
            // 
            this.label_ShipTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label_ShipTo.Location = new System.Drawing.Point(478, 98);
            this.label_ShipTo.Name = "label_ShipTo";
            this.label_ShipTo.Size = new System.Drawing.Size(189, 79);
            this.label_ShipTo.TabIndex = 11;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(475, 188);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(36, 13);
            this.label9.TabIndex = 12;
            this.label9.Text = "Terms";
            // 
            // comboBox_Terms
            // 
            this.comboBox_Terms.FormattingEnabled = true;
            this.comboBox_Terms.Location = new System.Drawing.Point(516, 180);
            this.comboBox_Terms.Name = "comboBox_Terms";
            this.comboBox_Terms.Size = new System.Drawing.Size(148, 21);
            this.comboBox_Terms.TabIndex = 13;
            // 
            // label11
            // 
            this.label11.Location = new System.Drawing.Point(15, 210);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(147, 14);
            this.label11.TabIndex = 15;
            this.label11.Text = "Item";
            // 
            // label12
            // 
            this.label12.Location = new System.Drawing.Point(160, 210);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(276, 14);
            this.label12.TabIndex = 16;
            this.label12.Text = "Description";
            // 
            // label13
            // 
            this.label13.Location = new System.Drawing.Point(496, 210);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(61, 14);
            this.label13.TabIndex = 17;
            this.label13.Text = "Price each";
            // 
            // label14
            // 
            this.label14.Location = new System.Drawing.Point(557, 210);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(49, 14);
            this.label14.TabIndex = 18;
            this.label14.Text = "Amount";
            // 
            // label15
            // 
            this.label15.Location = new System.Drawing.Point(606, 210);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(61, 14);
            this.label15.TabIndex = 19;
            this.label15.Text = "Tax (Y/N)";
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(442, 210);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(48, 14);
            this.label10.TabIndex = 20;
            this.label10.Text = "Quantity";
            // 
            // comboBox_Item1
            // 
            this.comboBox_Item1.FormattingEnabled = true;
            this.comboBox_Item1.Location = new System.Drawing.Point(14, 231);
            this.comboBox_Item1.Name = "comboBox_Item1";
            this.comboBox_Item1.Size = new System.Drawing.Size(148, 21);
            this.comboBox_Item1.TabIndex = 21;
            this.comboBox_Item1.SelectedValueChanged += new System.EventHandler(this.comboBox_Item1_SelectedValueChanged);
            // 
            // textBox_Desc1
            // 
            this.textBox_Desc1.Location = new System.Drawing.Point(160, 231);
            this.textBox_Desc1.Name = "textBox_Desc1";
            this.textBox_Desc1.Size = new System.Drawing.Size(276, 20);
            this.textBox_Desc1.TabIndex = 22;
            this.textBox_Desc1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Qty1
            // 
            this.textBox_Qty1.Location = new System.Drawing.Point(442, 232);
            this.textBox_Qty1.Name = "textBox_Qty1";
            this.textBox_Qty1.Size = new System.Drawing.Size(48, 20);
            this.textBox_Qty1.TabIndex = 23;
            this.textBox_Qty1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Qty1.TextChanged += new System.EventHandler(this.textBox_Qty1_TextChanged);
            // 
            // textBox_Price1
            // 
            this.textBox_Price1.Location = new System.Drawing.Point(496, 233);
            this.textBox_Price1.Name = "textBox_Price1";
            this.textBox_Price1.Size = new System.Drawing.Size(61, 20);
            this.textBox_Price1.TabIndex = 24;
            this.textBox_Price1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Price1.TextChanged += new System.EventHandler(this.textBox_Price1_TextChanged);
            // 
            // textBox_Amount1
            // 
            this.textBox_Amount1.Location = new System.Drawing.Point(557, 233);
            this.textBox_Amount1.Name = "textBox_Amount1";
            this.textBox_Amount1.Size = new System.Drawing.Size(49, 20);
            this.textBox_Amount1.TabIndex = 25;
            this.textBox_Amount1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Amount2
            // 
            this.textBox_Amount2.Location = new System.Drawing.Point(557, 259);
            this.textBox_Amount2.Name = "textBox_Amount2";
            this.textBox_Amount2.Size = new System.Drawing.Size(49, 20);
            this.textBox_Amount2.TabIndex = 31;
            this.textBox_Amount2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Price2
            // 
            this.textBox_Price2.Location = new System.Drawing.Point(496, 259);
            this.textBox_Price2.Name = "textBox_Price2";
            this.textBox_Price2.Size = new System.Drawing.Size(61, 20);
            this.textBox_Price2.TabIndex = 30;
            this.textBox_Price2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Price2.TextChanged += new System.EventHandler(this.textBox_Price2_TextChanged);
            // 
            // textBox_Qty2
            // 
            this.textBox_Qty2.Location = new System.Drawing.Point(442, 258);
            this.textBox_Qty2.Name = "textBox_Qty2";
            this.textBox_Qty2.Size = new System.Drawing.Size(48, 20);
            this.textBox_Qty2.TabIndex = 29;
            this.textBox_Qty2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Qty2.TextChanged += new System.EventHandler(this.textBox_Qty2_TextChanged);
            // 
            // textBox_Desc2
            // 
            this.textBox_Desc2.Location = new System.Drawing.Point(160, 257);
            this.textBox_Desc2.Name = "textBox_Desc2";
            this.textBox_Desc2.Size = new System.Drawing.Size(276, 20);
            this.textBox_Desc2.TabIndex = 28;
            this.textBox_Desc2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // comboBox_Item2
            // 
            this.comboBox_Item2.FormattingEnabled = true;
            this.comboBox_Item2.Location = new System.Drawing.Point(14, 257);
            this.comboBox_Item2.Name = "comboBox_Item2";
            this.comboBox_Item2.Size = new System.Drawing.Size(148, 21);
            this.comboBox_Item2.TabIndex = 27;
            this.comboBox_Item2.SelectedValueChanged += new System.EventHandler(this.comboBox_Item2_SelectedValueChanged);
            // 
            // textBox_Amount3
            // 
            this.textBox_Amount3.Location = new System.Drawing.Point(557, 287);
            this.textBox_Amount3.Name = "textBox_Amount3";
            this.textBox_Amount3.Size = new System.Drawing.Size(49, 20);
            this.textBox_Amount3.TabIndex = 37;
            this.textBox_Amount3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Price3
            // 
            this.textBox_Price3.Location = new System.Drawing.Point(496, 287);
            this.textBox_Price3.Name = "textBox_Price3";
            this.textBox_Price3.Size = new System.Drawing.Size(61, 20);
            this.textBox_Price3.TabIndex = 36;
            this.textBox_Price3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Price3.TextChanged += new System.EventHandler(this.textBox_Price3_TextChanged);
            // 
            // textBox_Qty3
            // 
            this.textBox_Qty3.Location = new System.Drawing.Point(442, 286);
            this.textBox_Qty3.Name = "textBox_Qty3";
            this.textBox_Qty3.Size = new System.Drawing.Size(48, 20);
            this.textBox_Qty3.TabIndex = 35;
            this.textBox_Qty3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Qty3.TextChanged += new System.EventHandler(this.textBox_Qty3_TextChanged);
            // 
            // textBox_Desc3
            // 
            this.textBox_Desc3.Location = new System.Drawing.Point(160, 285);
            this.textBox_Desc3.Name = "textBox_Desc3";
            this.textBox_Desc3.Size = new System.Drawing.Size(276, 20);
            this.textBox_Desc3.TabIndex = 34;
            this.textBox_Desc3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // comboBox_Item3
            // 
            this.comboBox_Item3.FormattingEnabled = true;
            this.comboBox_Item3.Location = new System.Drawing.Point(14, 285);
            this.comboBox_Item3.Name = "comboBox_Item3";
            this.comboBox_Item3.Size = new System.Drawing.Size(148, 21);
            this.comboBox_Item3.TabIndex = 33;
            this.comboBox_Item3.SelectedValueChanged += new System.EventHandler(this.comboBox_Item3_SelectedValueChanged);
            // 
            // textBox_Amount4
            // 
            this.textBox_Amount4.Location = new System.Drawing.Point(557, 315);
            this.textBox_Amount4.Name = "textBox_Amount4";
            this.textBox_Amount4.Size = new System.Drawing.Size(49, 20);
            this.textBox_Amount4.TabIndex = 43;
            this.textBox_Amount4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Price4
            // 
            this.textBox_Price4.Location = new System.Drawing.Point(496, 315);
            this.textBox_Price4.Name = "textBox_Price4";
            this.textBox_Price4.Size = new System.Drawing.Size(61, 20);
            this.textBox_Price4.TabIndex = 42;
            this.textBox_Price4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Price4.TextChanged += new System.EventHandler(this.textBox_Price4_TextChanged);
            // 
            // textBox_Qty4
            // 
            this.textBox_Qty4.Location = new System.Drawing.Point(442, 314);
            this.textBox_Qty4.Name = "textBox_Qty4";
            this.textBox_Qty4.Size = new System.Drawing.Size(48, 20);
            this.textBox_Qty4.TabIndex = 41;
            this.textBox_Qty4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Qty4.TextChanged += new System.EventHandler(this.textBox_Qty4_TextChanged);
            // 
            // textBox_Desc4
            // 
            this.textBox_Desc4.Location = new System.Drawing.Point(160, 313);
            this.textBox_Desc4.Name = "textBox_Desc4";
            this.textBox_Desc4.Size = new System.Drawing.Size(276, 20);
            this.textBox_Desc4.TabIndex = 40;
            this.textBox_Desc4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // comboBox_Item4
            // 
            this.comboBox_Item4.FormattingEnabled = true;
            this.comboBox_Item4.Location = new System.Drawing.Point(14, 313);
            this.comboBox_Item4.Name = "comboBox_Item4";
            this.comboBox_Item4.Size = new System.Drawing.Size(148, 21);
            this.comboBox_Item4.TabIndex = 39;
            this.comboBox_Item4.SelectedValueChanged += new System.EventHandler(this.comboBox_Item4_SelectedValueChanged);
            // 
            // textBox_Amount5
            // 
            this.textBox_Amount5.Location = new System.Drawing.Point(557, 341);
            this.textBox_Amount5.Name = "textBox_Amount5";
            this.textBox_Amount5.Size = new System.Drawing.Size(49, 20);
            this.textBox_Amount5.TabIndex = 49;
            this.textBox_Amount5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_Price5
            // 
            this.textBox_Price5.Location = new System.Drawing.Point(496, 341);
            this.textBox_Price5.Name = "textBox_Price5";
            this.textBox_Price5.Size = new System.Drawing.Size(61, 20);
            this.textBox_Price5.TabIndex = 48;
            this.textBox_Price5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Price5.TextChanged += new System.EventHandler(this.textBox_Price5_TextChanged);
            // 
            // textBox_Qty5
            // 
            this.textBox_Qty5.Location = new System.Drawing.Point(442, 340);
            this.textBox_Qty5.Name = "textBox_Qty5";
            this.textBox_Qty5.Size = new System.Drawing.Size(48, 20);
            this.textBox_Qty5.TabIndex = 47;
            this.textBox_Qty5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Qty5.TextChanged += new System.EventHandler(this.textBox_Qty5_TextChanged);
            // 
            // textBox_Desc5
            // 
            this.textBox_Desc5.Location = new System.Drawing.Point(160, 339);
            this.textBox_Desc5.Name = "textBox_Desc5";
            this.textBox_Desc5.Size = new System.Drawing.Size(276, 20);
            this.textBox_Desc5.TabIndex = 46;
            this.textBox_Desc5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // comboBox_Item5
            // 
            this.comboBox_Item5.FormattingEnabled = true;
            this.comboBox_Item5.Location = new System.Drawing.Point(14, 339);
            this.comboBox_Item5.Name = "comboBox_Item5";
            this.comboBox_Item5.Size = new System.Drawing.Size(148, 21);
            this.comboBox_Item5.TabIndex = 45;
            this.comboBox_Item5.SelectedValueChanged += new System.EventHandler(this.comboBox_Item5_SelectedValueChanged);
            // 
            // comboBox_Tax1
            // 
            this.comboBox_Tax1.FormattingEnabled = true;
            this.comboBox_Tax1.Location = new System.Drawing.Point(612, 233);
            this.comboBox_Tax1.Name = "comboBox_Tax1";
            this.comboBox_Tax1.Size = new System.Drawing.Size(55, 21);
            this.comboBox_Tax1.TabIndex = 51;
            // 
            // comboBox_Tax2
            // 
            this.comboBox_Tax2.FormattingEnabled = true;
            this.comboBox_Tax2.Location = new System.Drawing.Point(612, 260);
            this.comboBox_Tax2.Name = "comboBox_Tax2";
            this.comboBox_Tax2.Size = new System.Drawing.Size(55, 21);
            this.comboBox_Tax2.TabIndex = 52;
            // 
            // comboBox_Tax3
            // 
            this.comboBox_Tax3.FormattingEnabled = true;
            this.comboBox_Tax3.Location = new System.Drawing.Point(612, 286);
            this.comboBox_Tax3.Name = "comboBox_Tax3";
            this.comboBox_Tax3.Size = new System.Drawing.Size(55, 21);
            this.comboBox_Tax3.TabIndex = 53;
            // 
            // comboBox_Tax4
            // 
            this.comboBox_Tax4.FormattingEnabled = true;
            this.comboBox_Tax4.Location = new System.Drawing.Point(612, 315);
            this.comboBox_Tax4.Name = "comboBox_Tax4";
            this.comboBox_Tax4.Size = new System.Drawing.Size(55, 21);
            this.comboBox_Tax4.TabIndex = 54;
            // 
            // comboBox_Tax5
            // 
            this.comboBox_Tax5.FormattingEnabled = true;
            this.comboBox_Tax5.Location = new System.Drawing.Point(611, 342);
            this.comboBox_Tax5.Name = "comboBox_Tax5";
            this.comboBox_Tax5.Size = new System.Drawing.Size(55, 21);
            this.comboBox_Tax5.TabIndex = 55;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 404);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(97, 13);
            this.label6.TabIndex = 56;
            this.label6.Text = "Customer Message";
            // 
            // comboBox_CustomerMessage
            // 
            this.comboBox_CustomerMessage.FormattingEnabled = true;
            this.comboBox_CustomerMessage.Location = new System.Drawing.Point(16, 420);
            this.comboBox_CustomerMessage.Name = "comboBox_CustomerMessage";
            this.comboBox_CustomerMessage.Size = new System.Drawing.Size(334, 21);
            this.comboBox_CustomerMessage.TabIndex = 57;
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(459, 375);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(31, 13);
            this.label16.TabIndex = 60;
            this.label16.Text = "Total";
            // 
            // textBox_Total
            // 
            this.textBox_Total.Location = new System.Drawing.Point(496, 367);
            this.textBox_Total.Name = "textBox_Total";
            this.textBox_Total.Size = new System.Drawing.Size(110, 20);
            this.textBox_Total.TabIndex = 61;
            this.textBox_Total.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_Total.TextChanged += new System.EventHandler(this.textBox_Total_TextChanged);
            // 
            // label1_Currency
            // 
            this.label1_Currency.AutoSize = true;
            this.label1_Currency.Location = new System.Drawing.Point(612, 374);
            this.label1_Currency.Name = "label1_Currency";
            this.label1_Currency.Size = new System.Drawing.Size(49, 13);
            this.label1_Currency.TabIndex = 62;
            this.label1_Currency.Text = "Currency";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(459, 397);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(84, 13);
            this.label17.TabIndex = 63;
            this.label17.Text = "Exchange Rate:";
            this.label17.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // textBox_ExchangeRate
            // 
            this.textBox_ExchangeRate.Location = new System.Drawing.Point(549, 392);
            this.textBox_ExchangeRate.Name = "textBox_ExchangeRate";
            this.textBox_ExchangeRate.Size = new System.Drawing.Size(57, 20);
            this.textBox_ExchangeRate.TabIndex = 64;
            this.textBox_ExchangeRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_ExchangeRate.TextChanged += new System.EventHandler(this.textBox_ExchangeRate_TextChanged);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(421, 424);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(69, 13);
            this.label18.TabIndex = 66;
            this.label18.Text = "Balance Due";
            // 
            // textBox_BalanceDue
            // 
            this.textBox_BalanceDue.Location = new System.Drawing.Point(496, 421);
            this.textBox_BalanceDue.Name = "textBox_BalanceDue";
            this.textBox_BalanceDue.Size = new System.Drawing.Size(110, 20);
            this.textBox_BalanceDue.TabIndex = 67;
            this.textBox_BalanceDue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label3_Currency
            // 
            this.label3_Currency.AutoSize = true;
            this.label3_Currency.Location = new System.Drawing.Point(612, 424);
            this.label3_Currency.Name = "label3_Currency";
            this.label3_Currency.Size = new System.Drawing.Size(49, 13);
            this.label3_Currency.TabIndex = 68;
            this.label3_Currency.Text = "Currency";
            // 
            // button_Save
            // 
            this.button_Save.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_Save.Location = new System.Drawing.Point(499, 467);
            this.button_Save.Name = "button_Save";
            this.button_Save.Size = new System.Drawing.Size(161, 33);
            this.button_Save.TabIndex = 69;
            this.button_Save.Text = "Save";
            this.button_Save.UseVisualStyleBackColor = true;
            this.button_Save.Click += new System.EventHandler(this.button_Save_Click);
            // 
            // label4_Currency
            // 
            this.label4_Currency.AutoSize = true;
            this.label4_Currency.Location = new System.Drawing.Point(612, 398);
            this.label4_Currency.Name = "label4_Currency";
            this.label4_Currency.Size = new System.Drawing.Size(49, 13);
            this.label4_Currency.TabIndex = 70;
            this.label4_Currency.Text = "Currency";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(677, 510);
            this.Controls.Add(this.label4_Currency);
            this.Controls.Add(this.button_Save);
            this.Controls.Add(this.label3_Currency);
            this.Controls.Add(this.textBox_BalanceDue);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.textBox_ExchangeRate);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label1_Currency);
            this.Controls.Add(this.textBox_Total);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.comboBox_CustomerMessage);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.comboBox_Tax5);
            this.Controls.Add(this.comboBox_Tax4);
            this.Controls.Add(this.comboBox_Tax3);
            this.Controls.Add(this.comboBox_Tax2);
            this.Controls.Add(this.comboBox_Tax1);
            this.Controls.Add(this.textBox_Amount5);
            this.Controls.Add(this.textBox_Price5);
            this.Controls.Add(this.textBox_Qty5);
            this.Controls.Add(this.textBox_Desc5);
            this.Controls.Add(this.comboBox_Item5);
            this.Controls.Add(this.textBox_Amount4);
            this.Controls.Add(this.textBox_Price4);
            this.Controls.Add(this.textBox_Qty4);
            this.Controls.Add(this.textBox_Desc4);
            this.Controls.Add(this.comboBox_Item4);
            this.Controls.Add(this.textBox_Amount3);
            this.Controls.Add(this.textBox_Price3);
            this.Controls.Add(this.textBox_Qty3);
            this.Controls.Add(this.textBox_Desc3);
            this.Controls.Add(this.comboBox_Item3);
            this.Controls.Add(this.textBox_Amount2);
            this.Controls.Add(this.textBox_Price2);
            this.Controls.Add(this.textBox_Qty2);
            this.Controls.Add(this.textBox_Desc2);
            this.Controls.Add(this.comboBox_Item2);
            this.Controls.Add(this.textBox_Amount1);
            this.Controls.Add(this.textBox_Price1);
            this.Controls.Add(this.textBox_Qty1);
            this.Controls.Add(this.textBox_Desc1);
            this.Controls.Add(this.comboBox_Item1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.comboBox_Terms);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label_ShipTo);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label_BillTo);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox_RefNumber);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox_Customer);
            this.Controls.Add(this.labelLoadStatus);
            this.Name = "Form1";
            this.Text = "Multicurrency Invoice";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelLoadStatus;
        private System.Windows.Forms.ComboBox comboBox_Customer;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_RefNumber;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label_BillTo;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label_ShipTo;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBox_Terms;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox comboBox_Item1;
        private System.Windows.Forms.TextBox textBox_Desc1;
        private System.Windows.Forms.TextBox textBox_Qty1;
        private System.Windows.Forms.TextBox textBox_Price1;
        private System.Windows.Forms.TextBox textBox_Amount1;
        private System.Windows.Forms.TextBox textBox_Amount2;
        private System.Windows.Forms.TextBox textBox_Price2;
        private System.Windows.Forms.TextBox textBox_Qty2;
        private System.Windows.Forms.TextBox textBox_Desc2;
        private System.Windows.Forms.ComboBox comboBox_Item2;
        private System.Windows.Forms.TextBox textBox_Amount3;
        private System.Windows.Forms.TextBox textBox_Price3;
        private System.Windows.Forms.TextBox textBox_Qty3;
        private System.Windows.Forms.TextBox textBox_Desc3;
        private System.Windows.Forms.ComboBox comboBox_Item3;
        private System.Windows.Forms.TextBox textBox_Amount4;
        private System.Windows.Forms.TextBox textBox_Price4;
        private System.Windows.Forms.TextBox textBox_Qty4;
        private System.Windows.Forms.TextBox textBox_Desc4;
        private System.Windows.Forms.ComboBox comboBox_Item4;
        private System.Windows.Forms.TextBox textBox_Amount5;
        private System.Windows.Forms.TextBox textBox_Price5;
        private System.Windows.Forms.TextBox textBox_Qty5;
        private System.Windows.Forms.TextBox textBox_Desc5;
        private System.Windows.Forms.ComboBox comboBox_Item5;
        private System.Windows.Forms.ComboBox comboBox_Tax1;
        private System.Windows.Forms.ComboBox comboBox_Tax2;
        private System.Windows.Forms.ComboBox comboBox_Tax3;
        private System.Windows.Forms.ComboBox comboBox_Tax4;
        private System.Windows.Forms.ComboBox comboBox_Tax5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox_CustomerMessage;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox textBox_Total;
        private System.Windows.Forms.Label label1_Currency;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox textBox_ExchangeRate;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox textBox_BalanceDue;
        private System.Windows.Forms.Label label3_Currency;
        private System.Windows.Forms.Button button_Save;
        private System.Windows.Forms.Label label4_Currency;
    }
}

