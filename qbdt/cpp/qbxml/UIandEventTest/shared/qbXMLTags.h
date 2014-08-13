/*---------------------------------------------------------------------------
 * FILE: qbXMLTags.h
 *
 * Description:
 * Tag definition
 *
 * Created On: 11/11/2001
 *
 *
 * Copyright © 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#ifndef QBXMLTAGS_H
#define QBXMLTAGS_H

#define QBXML_TAG "QBXML"

// Aggregated Tags

#define QBXMLMSGSRQ_TAG         "QBXMLMsgsRq"
#define INVOICEQUERYRQ_TAG      "InvoiceQueryRq"
#define ESTIMATEQUERYRQ_TAG     "EstimateQueryRq"
#define CREDITMEMOQUERYRQ_TAG   "CreditMemoQueryRq"
#define PURCHASEORDERQUERYRQ_TAG  "PurchaseOrderQueryRq"
#define CUSTOMERQUERY_TAG       "CustomerQueryRq"
#define EMPLOYEEQUERYRQRQ_TAG   "EmployeeQueryRq"
#define VENDORQUERYRQ_TAG       "VendorQueryRq"
#define OTHERNAMEQUERYRQ_TAG    "OtherNameQueryRq"
#define QBXMLSUBMSGSRQ_TAG      "QBXMLSubscriptionMsgsRq"
#define QBXMLSUBMSGSRS_TAG      "QBXMLSubscriptionMsgsRs"
#define QBXMLMSGSRS_TAG         "QBXMLMsgsRs"
#define QBXMLSUBMSGSRS_TAG      "QBXMLSubscriptionMsgsRs"
#define SUBSCRIPTIONDELRQ_TAG	"SubscriptionDelRq"
#define SUBSCRIPTIONTYPE_TAG	"SubscriptionType"
#define UIEXTSUBADDRQ_TAG		"UIExtensionSubscriptionAddRq"
#define UIEXTSUBADD_TAG			"UIExtensionSubscriptionAdd"
#define UIEXTSUBQUERYRQ_TAG		"UIExtensionSubscriptionQueryRq"
#define UIEVENTSUBADDRQ_TAG	    "UIEventSubscriptionAddRq"
#define UIEVENTSUBADD_TAG		"UIEventSubscriptionAdd"
#define UIEVENTSUBQUERYRQ_TAG	"UIEventSubscriptionQueryRq"
#define DATAEVENTSUBADDRQ_TAG	"DataEventSubscriptionAddRq"
#define DATAEVENTSUBADD_TAG		"DataEventSubscriptionAdd"
#define DATAEVENTSUBQUERYRQ_TAG	"DataEventSubscriptionQueryRq"
#define DATAEVENTRECOVERYINFODELRQ_TAG "DataEventRecoveryInfoDelRq"
#define SUBSCRIBERID_TAG		"SubscriberID"
#define COMCALLBACKINFO_TAG		"COMCallbackInfo"
#define CLSID_TAG				"CLSID"
#define TRACKLOSTEVENTS_TAG     "TrackLostEvents"
#define DELIVERYPOLICY_TAG      "DeliveryPolicy"
#define APPNAME_TAG				"AppName"
#define COMPANYFILEEVENTSUB_TAG	"CompanyFileEventSubscription"
#define COMPANYFILEEVENTOP_TAG	"CompanyFileEventOperation"
#define LISTEVENTSUB_TAG		"ListEventSubscription"
#define LISTEVENTTYPE_TAG       "ListEventType"
#define LISTEVENTOP_TAG         "ListEventOperation"
#define TXNEVENTSUB_TAG	  	    "TxnEventSubscription"
#define TXNEVENTTYPE_TAG        "TxnEventType"
#define TXNEVENTOP_TAG          "TxnEventOperation"
#define MENUEXTSUB_TAG			"MenuExtensionSubscription"
#define ADDTOMENU_TAG			"AddToMenu"
#define SUBMENU_TAG				"Submenu"
#define MENUITEM_TAG			"MenuItem"
#define MENUTEXT_TAG			"MenuText"
#define EVENTTAG_TAG			"EventTag"
#define DISPLAYCONDITION_TAG	"DisplayCondition"
#define VISIBLEIF_TAG			"VisibleIf"
#define VISIBLEIFNOT_TAG		"VisibleIfNot"
#define ENABLEDIF_TAG			"EnabledIf"
#define ENABLEDIFNOT_TAG		"EnabledIfNot"
#define DATAEVENTRECOVERYTIME_TAG  "DataEventRecoveryTime"
#define COMPANYFILEPATH_TAG     "CompanyFilePath"

// Attributes

#define ATTR_ONERROR_TAG          "onError"
#define ATTR_REQUESTID_TAG        "requestID"
#define ATTR_STATUSCODE_TAG       "statusCode"
#define ATTR_STATUSMESSAGE_TAG    "statusMessage"
#define ATTR_STATUSSEVERITY_TAG   "statusSeverity"
#define ATTR_TOKEN_TAG            "token"
#define ATTR_TYPE_TAG             "type"
#define ATTR_VERSION_TAG          "version"

// Enumerations
#define SUBSCRIPTION_TYPE_DATAEVENT		"Data"
#define SUBSCRIPTION_TYPE_UIEVENT		"UI"
#define SUBSCRIPTION_TYPE_UIEXTENSION	"UIExtension"
#define EVENT_OPERATION_ADD             "Add"
#define EVENT_OPERATION_DELETE          "Delete"
#define EVENT_OPERATION_MODIFY          "Modify"
#define EVENT_OPERATION_MERGE           "Merge"
#define DELIVERY_POLICY_ALWAYS          "DeliverAlways"
#define DELIVERY_POLICY_ONLY_IF_RUNNING "DeliverOnlyIfRunning"
#define COMPANY_FILE_EVENT_OPEN         "Open"
#define COMPANY_FILE_EVENT_CLOSE        "Close"

// Values
#define SUBSCRIBER_ID		"{70DC83AB-2EC1-478F-9204-C37485AD8E3C}"
#define HANDLER_APP_NAME	"QBSDK Event and UI Sample"
#define HANDLER_CLSID		"{70DC83AB-2EC1-478F-9204-C37485AD8E3C}"

// Elements

#endif QBXMLTAGS_H
