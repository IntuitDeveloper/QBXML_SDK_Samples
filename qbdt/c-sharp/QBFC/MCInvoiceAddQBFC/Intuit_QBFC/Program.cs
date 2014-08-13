// Copyright (c) 2007-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using System.Collections.Generic;
using System.Windows.Forms;
using MCInvoiceAddQBFC.Session_Framework;

namespace MCInvoiceAddQBFC
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                SessionManager sessionMgr = SessionManager.getInstance();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
                Logger.getInstance().logInfo("Application is closing normally");
            }
            catch (Exception e)
            {
                Logger log = Logger.getInstance();
                log.logCritical("Program.Main", e.Message);
                log.logCritical("Program.Main", "Application Terminating");
                MessageBox.Show("An error requires the application to close:\n" +
                                "Logging File: " + Logger.getInstance().LogLevel.ToString() + "\n" +
                                "Log File: " + log.LogFile + "\n" +
                                "Exception Caught: \n" + e.Message);
            }
        }
    }
}