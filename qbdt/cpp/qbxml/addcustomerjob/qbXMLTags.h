/*---------------------------------------------------------------------------
 * FILE: qbXMLTags.h
 *
 * Description:
 * Tag definition
 *
 * Created On: 11/11/2001
 * Updated to SDK 2.0: 08/08/2002
 *
 * Copyright © 2002-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#ifndef QBXMLTAGS_H
#define QBXMLTAGS_H

#define QBXML_TAG "QBXML"

// Aggregated Tags

#define QBXMLMSGSRQ_TAG           "QBXMLMsgsRq"
#define QBXMLMSGSRS_TAG           "QBXMLMsgsRs"
#define CUSTOMERADD_TAG           "CustomerAdd"
#define CUSTOMERADDRQ_TAG         "CustomerAddRq"
#define CUSTOMERADDRS_TAG         "CustomerAddRs"
#define CUSTOMERMODRQ_TAG         "CustomerModRq"
#define CUSTOMERMODRS_TAG         "CustomerModRs"
#define CUSTOMERQUERYRQ_TAG       "CustomerQueryRq"
#define CUSTOMERQUERYRS_TAG       "CustomerQueryRs"
#define CUSTOMERRET_TAG           "CustomerRet"
#define CUSTOMERTYPEQUERYRQ_TAG   "CustomerTypeQueryRq"
#define CUSTOMERTYPEQUERYRS_TAG   "CustomerTypeQueryRs"
#define CUSTOMERTYPEADD_TAG       "CustomerTypeAdd"
#define CUSTOMERTYPEADDRQ_TAG     "CustomerTypeAddRq"
#define CUSTOMERTYPEADDRS_TAG     "CustomerTypeAddRs"
#define CUSTOMERTYPEREF_TAG       "CustomerTypeRef"
#define CUSTOMERTYPERET_TAG       "CustomerTypeRet"
#define PARENTREF_TAG             "ParentRef"

#define EDITSEQUENCE_TAG          "EditSequence"
#define ISACTIVE_TAG              "IsActive"
#define TRANSACTIONID_TAG         "TransactionID"


// Attributes

#define ATTR_ONERROR_TAG          "onError"
#define ATTR_REQUESTID_TAG        "requestID"
#define ATTR_STATUSCODE_TAG       "statusCode"
#define ATTR_STATUSMESSAGE_TAG    "statusMessage"
#define ATTR_STATUSSEVERITY_TAG   "statusSeverity"
#define ATTR_TOKEN_TAG            "token"
#define ATTR_TRNUID_TAG           "trnUID"
#define ATTR_TYPE_TAG             "type"
#define ATTR_VERSION_TAG          "version"

// Elements

#define BALANCE_TAG               "Balance"
#define COMPANY_TAG               "Company"
#define COMPANYID_TAG             "CompanyID"
#define COMPANYNAME_TAG           "CompanyName"
#define CONTACT_TAG               "Contact"
#define COUNTRY_TAG               "Country"
#define CREATEDATE_TAG            "CreateDate"
#define CREATEUSER_TAG            "CreateUser"
#define CUSTOMERID_TAG            "CustomerID"
#define CUSTOMERJOBID_TAG         "CustomerJobID"
#define CUSTOMERNAME_TAG          "CustomerName"
#define EMAIL_TAG                 "Email"
#define FIRSTNAME_TAG             "FirstName"
#define FULLNAME_TAG              "FullName"
#define INVOICE_TAG               "Invoice"
#define INVOICEAMT_TAG            "InvoiceAmt"
#define INVOICEDAMOUNT_TAG        "InvoicedAmount"
#define INVOICEDATA_TAG           "InvoiceData"
#define INVOICEDETAIL_TAG         "InvoiceDetail"
#define INVOICEID_TAG             "InvoiceID"
#define INVOICEITEM_TAG           "InvoiceItem"
#define INVOICELINE_TAG           "InvoiceLine"
#define INVOICENUM_TAG            "InvoiceNum"
#define LASTNAME_TAG              "LastName"
#define LISTID_TAG                "ListID"
#define MIDDLENAME_TAG            "MiddleName"
#define NAME_TAG                  "Name"
#define PHONE_TAG                 "Phone"
#define STATE_TAG                 "State"
#define SUBLEVEL_TAG              "Sublevel"
#define SUBTOTAL_TAG              "Subtotal"
#define SUGGESTEDDISCOUNT_TAG     "SuggestedDiscount"
#define SUGGESTEDDISCOUNTDATA_TAG "SuggestedDiscountData"
#define TIMECREATED_TAG           "TimeCreated"
#define TIMEMODIFIED_TAG          "TimeModified"
#define TOTALBALANCE_TAG          "TotalBalance"
#define TOTALAMT_TAG              "TotalAmt"
#define TOTALPAYMENT_TAG          "TotalPayment"
#define TRANID_TAG                "TranID"

#endif QBXMLTAGS_H
