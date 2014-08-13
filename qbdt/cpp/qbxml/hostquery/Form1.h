#pragma once

/*---------------------------------------------------------------------------
 * QBXML/VC++ HostQuery sample using Visual Studio 2005
 *
 * Description:
 * This is a very simple vc++/QBXML sample written using Visual Studio 2005. 
 * It has been created just to demonstrate how to use vc++ to communicate 
 * with QuickBooks via QBXMLRP2.
 * 
 * Created On: 05/09/2007
 * Uses: QBXMLRP2
 *
 * Copyright © 2007-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *---------------------------------------------------------------------------*/

namespace hostquery {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace Interop::QBXMLRP2;
	using namespace System::Xml;

	/// <summary>
	/// Summary for Form1
	///
	/// WARNING: If you change the name of this class, you will need to change the
	///          'Resource File Name' property for the managed resource compiler tool
	///          associated with all .resx files this class depends on.  Otherwise,
	///          the designers will not be able to interact properly with localized
	///          resources associated with this form.
	/// </summary>
	public ref class Form1 : public System::Windows::Forms::Form
	{
	public:
		Form1(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~Form1()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::Label^  label1;
	protected: 

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// button1
			// 
			this->button1->Font = (gcnew System::Drawing::Font(L"Courier New", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(0)));
			this->button1->Location = System::Drawing::Point(47, 115);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(167, 49);
			this->button1->TabIndex = 0;
			this->button1->Text = L"Run Host Query";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &Form1::button1_Click);
			// 
			// label1
			// 
			this->label1->Font = (gcnew System::Drawing::Font(L"Courier New", 8.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, 
				static_cast<System::Byte>(0)));
			this->label1->Location = System::Drawing::Point(12, 9);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(247, 90);
			this->label1->TabIndex = 1;
			this->label1->Text = L"label1";
			// 
			// Form1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(276, 171);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->button1);
			this->Name = L"Form1";
			this->Text = L"Form1";
			this->Load += gcnew System::EventHandler(this, &Form1::Form1_Load);
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
				/*
				A note on the caret (^): -
				According to MSDN/C++ Language Reference, ^ declares a handle to an object on the managed heap.
				A handle to an object on the managed heap points to the "whole" object, and not to a member of the object.
				See gcnew for information on how to create an object on the managed heap.
				In Visual C++ 2002 and Visual C++ 2003, __gc * was used to declare an object on the managed heap. The ^ 
				replaces __gc * in the new syntax.  The common language runtime maintains a separate heap on which it 
				implements a precise, asynchronous, compacting garbage collection scheme. To work correctly, it must track 
				all storage locations that can point into this heap at runtime. ^ provides a handle through which the 
				garbage collector can track a reference to an object on the managed heap, thereby being able to update it 
				whenever that object is moved. Because regular C++ pointers (*) and references (&) cannot be tracked 
				precisely, a handle-to object declarator is used. Member selection through a handle (^) uses the 
				pointer-to-member operator (->). 
				*/
				String ^ticket;  
				String ^request=BuildRequestXML();
				String ^response;

				// Initialize request processor
				Interop::QBXMLRP2::IRequestProcessor4 ^rqPtr= gcnew Interop::QBXMLRP2::RequestProcessor2; 
				// OpenConnection
				rqPtr ->OpenConnection2("123", "IDN CPP HostQuery", Interop::QBXMLRP2::QBXMLRPConnectionType::localQBD);
				// BeginSession
				ticket=rqPtr ->BeginSession("", Interop::QBXMLRP2::QBFileMode::qbFileOpenDoNotCare);
				// Show request xml
				MessageBox::Show(request);
				// ProcessRequest
				response = rqPtr->ProcessRequest(ticket, request);
				// Show response xml
				MessageBox::Show(response);
				// EndSession
				rqPtr->EndSession (ticket);
				// CloseConnection
				rqPtr->CloseConnection();
				MessageBox::Show("Host query completed");
			 }

	private: String^ BuildRequestXML(){
				XmlDocument^ doc = gcnew XmlDocument;
				doc->AppendChild(doc->CreateXmlDeclaration("1.0", nullptr, nullptr));
				doc->AppendChild(doc->CreateProcessingInstruction("qbxml", "version=\"2.0\""));
				XmlElement^ qbXML = doc->CreateElement("QBXML");
				doc->AppendChild(qbXML);
				XmlElement^ qbXMLMsgsRq = doc->CreateElement("QBXMLMsgsRq");
				qbXML->AppendChild(qbXMLMsgsRq);
				qbXMLMsgsRq->SetAttribute("onError", "stopOnError");
				
				XmlElement^ HQRq = doc->CreateElement("HostQueryRq");
				qbXMLMsgsRq->AppendChild(HQRq);
				HQRq->SetAttribute("requestID", "1");
				String^ input=doc->OuterXml;
				return input;
			 }

	private: System::Void Form1_Load(System::Object^  sender, System::EventArgs^  e) {
				 label1->Text="This is a very simple vc++/QBXML sample written using Visual Studio 2005. ";
				 label1->Text=label1->Text+"It has been created just to demonstrate how to use vc++ to communicate ";
				 label1->Text=label1->Text+"with QuickBooks via QBXMLRP2.";
			 }
};
}

