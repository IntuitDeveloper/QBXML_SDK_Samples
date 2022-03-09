Option Strict Off
Option Explicit On
Friend Class Customer
    '-----------------------------------------------------------
    ' Class Module: Customer
    '
    ' Description:  Provides get/let methods to access customer data.
    '
    ' Created On: 11/09/2001
    ' Updated to SDK 2.0: 08/05/2002
    '
    ' Copyright © 2021-2022 Intuit Inc. All rights reserved.
    ' Use is subject to the terms specified at:
    '      http://developer.intuit.com/legal/devsite_tos.html
    '
    '----------------------------------------------------------


    Private m_ListID As String
	Private m_TimeCreated As String
	Private m_TimeModified As String
	Private m_EditSequence As String
	Private m_Name As String
	Private m_FullName As String
	Private m_IsActive As String
	Private m_Sublevel As String
	Private m_FirstName As String
	Private m_LastName As String
	Private m_CompanyName As String
	Private m_BillAddress_Addr1 As String
	Private m_BillAddress_Addr2 As String
	Private m_BillAddress_City As String
	Private m_BillAddress_State As String
	Private m_BillAddress_PostalCode As String
	Private m_Phone As String
	Private m_Email As String
	Private m_Contact As String
	
	Private m_CustomerTypeRef_ListID As String
	Private m_CustomerTypeRef_Residential As String
	
	Private m_TermsRef_ListID As String
	Private m_TermsRef_FullName As String
	
	Private m_Balance As String
	Private m_TotalBalance As String
	
	Private m_SalesTaxCodeRef_ListID As String
	Private m_SalesTaxCodeRef_FullName As String
	
	Private m_ItemSalesTaxRef_ListID As String
	Private m_ItemSalesTaxRef_FullName As String
	
	Private m_Status As String
	
	
	Public Property ListID() As String
		Get
			ListID = m_ListID
		End Get
		Set(ByVal Value As String)
			m_ListID = Value
		End Set
	End Property
	
	
	Public Property FullName() As String
		Get
			FullName = m_FullName
		End Get
		Set(ByVal Value As String)
			m_FullName = Value
		End Set
	End Property
	
	
	Public Property TimeCreated() As String
		Get
			TimeCreated = m_TimeCreated
		End Get
		Set(ByVal Value As String)
			m_TimeCreated = Value
		End Set
	End Property
	
	
	Public Property TimeModified() As String
		Get
			TimeModified = m_TimeModified
		End Get
		Set(ByVal Value As String)
			m_TimeModified = Value
		End Set
	End Property
	
	
	Public Property EditSequence() As String
		Get
			EditSequence = m_EditSequence
		End Get
		Set(ByVal Value As String)
			m_EditSequence = Value
		End Set
	End Property
	
	
	Public Property Name() As String
		Get
			Name = m_Name
		End Get
		Set(ByVal Value As String)
			m_Name = Value
		End Set
	End Property
	
	
	Public Property IsActive() As String
		Get
			IsActive = m_IsActive
		End Get
		Set(ByVal Value As String)
			m_IsActive = Value
		End Set
	End Property
	
	
	Public Property Sublevel() As String
		Get
			Sublevel = m_Sublevel
		End Get
		Set(ByVal Value As String)
			m_Sublevel = Value
		End Set
	End Property
	
	
	Public Property FirstName() As String
		Get
			FirstName = m_FirstName
		End Get
		Set(ByVal Value As String)
			m_FirstName = Value
		End Set
	End Property
	
	
	Public Property LastName() As String
		Get
			LastName = m_LastName
		End Get
		Set(ByVal Value As String)
			m_LastName = Value
		End Set
	End Property
	
	
	Public Property BillAddress_Addr1() As String
		Get
			BillAddress_Addr1 = m_BillAddress_Addr1
		End Get
		Set(ByVal Value As String)
			m_BillAddress_Addr1 = Value
		End Set
	End Property
	
	
	Public Property BillAddress_Addr2() As String
		Get
			BillAddress_Addr2 = m_BillAddress_Addr2
		End Get
		Set(ByVal Value As String)
			m_BillAddress_Addr2 = Value
		End Set
	End Property
	
	
	Public Property BillAddress_City() As String
		Get
			BillAddress_City = m_BillAddress_City
		End Get
		Set(ByVal Value As String)
			m_BillAddress_City = Value
		End Set
	End Property
	
	
	Public Property BillAddress_State() As String
		Get
			BillAddress_State = m_BillAddress_State
		End Get
		Set(ByVal Value As String)
			m_BillAddress_State = Value
		End Set
	End Property
	
	
	Public Property BillAddress_PostalCode() As String
		Get
			BillAddress_PostalCode = m_BillAddress_PostalCode
		End Get
		Set(ByVal Value As String)
			m_BillAddress_PostalCode = Value
		End Set
	End Property
	
	
	Public Property Phone() As String
		Get
			Phone = m_Phone
		End Get
		Set(ByVal Value As String)
			m_Phone = Value
		End Set
	End Property
	
	
	Public Property Email() As String
		Get
			Email = m_Email
		End Get
		Set(ByVal Value As String)
			m_Email = Value
		End Set
	End Property
	
	
	Public Property Contact() As String
		Get
			Contact = m_Contact
		End Get
		Set(ByVal Value As String)
			m_Contact = Value
		End Set
	End Property
	
	
	Public Property CustomerTypeRef_ListID() As String
		Get
			CustomerTypeRef_ListID = m_CustomerTypeRef_ListID
		End Get
		Set(ByVal Value As String)
			m_CustomerTypeRef_ListID = Value
		End Set
	End Property
	
	
	Public Property CustomerTypeRef_Residential() As String
		Get
			CustomerTypeRef_Residential = m_CustomerTypeRef_Residential
		End Get
		Set(ByVal Value As String)
			m_CustomerTypeRef_Residential = Value
		End Set
	End Property
	
	
	Public Property TermsRef_ListID() As String
		Get
			TermsRef_ListID = m_TermsRef_ListID
		End Get
		Set(ByVal Value As String)
			m_TermsRef_ListID = Value
		End Set
	End Property
	
	
	Public Property TermsRef_FullName() As String
		Get
			TermsRef_FullName = m_TermsRef_FullName
		End Get
		Set(ByVal Value As String)
			m_TermsRef_FullName = Value
		End Set
	End Property
	
	
	Public Property CompanyName() As String
		Get
			CompanyName = m_CompanyName
		End Get
		Set(ByVal Value As String)
			m_CompanyName = Value
		End Set
	End Property
	
	
	Public Property Balance() As String
		Get
			Balance = m_Balance
		End Get
		Set(ByVal Value As String)
			m_Balance = Value
		End Set
	End Property
	
	
	Public Property TotalBalance() As String
		Get
			TotalBalance = m_TotalBalance
		End Get
		Set(ByVal Value As String)
			m_TotalBalance = Value
		End Set
	End Property
	
	
	Public Property SalesTaxCodeRef_ListID() As String
		Get
			SalesTaxCodeRef_ListID = m_SalesTaxCodeRef_ListID
		End Get
		Set(ByVal Value As String)
			m_SalesTaxCodeRef_ListID = Value
		End Set
	End Property
	
	
	Public Property SalesTaxCodeRef_FullName() As String
		Get
			SalesTaxCodeRef_FullName = m_SalesTaxCodeRef_FullName
		End Get
		Set(ByVal Value As String)
			m_SalesTaxCodeRef_FullName = Value
		End Set
	End Property
	
	
	Public Property ItemSalesTaxRef_ListID() As String
		Get
			ItemSalesTaxRef_ListID = m_ItemSalesTaxRef_ListID
		End Get
		Set(ByVal Value As String)
			m_ItemSalesTaxRef_ListID = Value
		End Set
	End Property
	
	
	Public Property ItemSalesTaxRef_fullName() As String
		Get
			ItemSalesTaxRef_fullName = m_ItemSalesTaxRef_FullName
		End Get
		Set(ByVal Value As String)
			m_ItemSalesTaxRef_FullName = Value
		End Set
	End Property
End Class