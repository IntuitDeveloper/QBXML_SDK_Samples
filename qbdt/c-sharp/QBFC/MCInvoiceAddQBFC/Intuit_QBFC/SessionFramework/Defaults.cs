// Copyright (c) 2007-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using Interop.QBFC13;
using System.Runtime.InteropServices;


namespace MCInvoiceAddQBFC.Session_Framework
{
    static class Defaults
    {
        #region Connection and Session variables
        /// <summary>
        /// The default connection type used when creating a connection with QuickBooks.
        /// This may be manually overridden within the initalize(...) method's parameters.
        /// </summary>
        public const ENConnectionType CONNECTION_TYPE = ENConnectionType.ctLocalQBD;

        /// <summary>
        /// The default connection mode used when opening a session with QuickBooks.
        /// A different value can be manually provided within the SessionManager.beginSession(...)
        /// method.
        /// </summary>
        public const ENOpenMode SESSION_MODE = ENOpenMode.omDontCare;

        /// <summary>
        /// The Edition of QuickBooks for which this is designed to be used.
        /// </summary>
        public const ENEdition EDITION = ENEdition.edUS;

        /// <summary>
        /// The full path to the company file that should be used when creating a session
        /// with QuickBooks.  This is only mandatory when running in unattended mode.
        /// </summary>
        public const string QBFILE = "";
        #endregion

        #region Logging variables
        /// <summary>
        /// Product Version string
        /// </summary>
        public const string PRODUCT_VERSION = "2.0";

        /// <summary>
        /// The maxmimum size to which a log file can grow before it is renamed xxxOLD and
        /// a new file is created
        /// </summary>
        public const int MAX_LOG_FILESIZE = 100 * 1024;

        /// <summary>
        /// The default size of the log file
        /// </summary>
        public const int DEFAULT_LOG_FILESIZE = 50 * 1024;

        /// <summary>
        /// The default logging level
        /// </summary>
        public const ENLogLevel DEFAULT_LOG_LEVEL = ENLogLevel.CRITICAL;
        #endregion

        #region Automatically generated (during the wizard) variables
        /// <summary>
        /// The application name used when opening a connection with QuickBooks
        /// </summary>
        public const string APPNAME = "MCInvoiceAddQBFC";

        /// <summary>
        /// The Application ID, obtained from Intuit for you specific application, used
        /// when opening a connection with QuickBooks
        /// </summary>
        public const string APPID = "";

        /// <summary>
        /// The ID used to uniquely identify the QBFC application when subscribing to
        /// events within QuickBooks
        /// </summary>
        public static readonly Guid SUBSCRIBER_ID = new Guid("{6edde65c-99c4-4954-b4e6-4cd43fef596f}");

        /// <summary>
        /// the Registration Key for your application within the Microsoft Registry
        /// </summary>
        public const string REG_KEY = "Software\\IDN\\MCInvoiceAddQBFC";


        #endregion
    }

}
