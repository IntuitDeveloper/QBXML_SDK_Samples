// Copyright (c) 2007-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms;


namespace MCInvoiceAddQBFC.Session_Framework
{
    // Copyright (c) 2007-2013 by Intuit Inc.
    // All rights reserved
    // Usage governed by the QuickBooks SDK Developer's License Agreement

    /// <summary>
    /// Permissible Logging Request priorities
    /// </summary>
    public enum ENLogLevel
    {
        NONE,                // No logging of any type is done
        CRITICAL,            // Only Exception logging is undertaken
        ERROR,               // Basic logging of errors and XML traces
        VERBOSE,             // Full logging of any calls to the logging framework
    }

    /// <summary>
    /// Summary description for Logger.
    /// </summary>
    public class Logger
    {
        private string _productName = Defaults.APPNAME;
        private string _productVersion = Defaults.PRODUCT_VERSION;
        private static Logger _logger = null;
        private string _logFileName = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\Intuit\\QBSDK\\log\\" + Defaults.APPNAME.Replace(" ", "") + "Log.txt";
        private string _oldLogFileName = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\Intuit\\QBSDK\\log\\" + Defaults.APPNAME.Replace(" ", "") + "LogOld.txt";
        private int _maxLogFileSize = Defaults.MAX_LOG_FILESIZE;
        private ENLogLevel _logLevel = Defaults.DEFAULT_LOG_LEVEL;
        private StreamWriter _str;
        private const string MaxLogSizeKey = "MaxLogFileSize";
        private const string LogLevelKey = "LogLevel";

        #region Public Interface
        /// <summary>
        /// Sets the new logging level
        /// </summary>
        public ENLogLevel LogLevel
        {
            get
            {
                return _logLevel;
            }
            set
            {
                _logLevel = value;
                setLogLevel(_logLevel);
            }
        }

        /// <summary>
        /// Gets the log filename (Set is not available)
        /// </summary>
        public string LogFile
        {
            get
            {
                return _logFileName;
            }
        }

        /// <summary>
        /// Private as this is a singleton.  Access through the getLogger() method
        /// </summary>
        private Logger()
        {
            InitializeLog();
        }

        /// <summary>
        /// Obtains the logging object
        /// </summary>
        /// <returns>The logging object</returns>
        public static Logger getInstance()
        {
            if (_logger == null)
            {
                _logger = new Logger();
            }
            return _logger;
        }

        /// <summary>
        /// Clears the present log file
        /// </summary>
        public void clearLog()
        {
            if (File.Exists(_logFileName))
            {
                _str = File.CreateText(_logFileName);
                _str.Close();
            }
        }

        /// <summary>
        /// Initiates viewing of the log file
        /// </summary>
        public void viewLog()
        {
            try
            {
                System.Diagnostics.Process.Start(_logFileName);
            }
            catch (Exception eLog)
            {
                if (eLog.Message.IndexOf("The system cannot find the file specified") >= 0)
                {
                    try
                    {
                        StreamWriter SW = File.CreateText(_logFileName);
                        SW.WriteLine("Application Log");
                        SW.WriteLine("------------------");
                        SW.Close();
                        InitializeLog();
                        System.Diagnostics.Process.Start(_logFileName);
                    }
                    catch
                    {
                        MessageBox.Show("Unable to execute Logger.viewLog():\n" + eLog.Message);
                    }
                }
                else
                    MessageBox.Show("Unable to execute Logger.viewLog():\n" + eLog.Message);
            }
        }

        /// <summary>
        /// Attempts to log a critical message.  The actual logging may or may not occur as a 
        /// funciton of the current Log Level.
        /// NOTE: Please use the logCritical(method, text) method wherever possible
        /// </summary>
        /// <param name="logText">The message to log</param>
        public void logCritical(string logText)
        {
            logCritical("", logText);
        }

        /// <summary>
        /// Attempts to log a critical message.  The actual logging may or may not occur as a 
        /// funciton of the current Log Level.
        /// </summary>
        /// <param name="method">The location from which the log message was requested</param>
        /// <param name="logText">The message to log</param>
        public void logCritical(string method, string logText)
        {
            if (_logLevel == ENLogLevel.NONE)
                return;

            log("[CRITICAL]", method, logText);
        }

        /// <summary>
        /// Attempts to log an Error message.  The actual logging may or may not occur as a 
        /// funciton of the current Log Level.
        /// NOTE: Please use the logError(method, text) method wherever possible
        /// </summary>
        /// <param name="logText">The message to log</param>
        public void logError(string logText)
        {
            logError("", logText);
        }

        /// <summary>
        /// Attempts to log an Error message.  The actual logging may or may not occur as a 
        /// funciton of the current Log Level.
        /// </summary>
        /// <param name="method">The location from which the log message was requested</param>
        /// <param name="logText">The message to log</param>
        public void logError(string method, string logText)
        {
            if (_logLevel == ENLogLevel.NONE || _logLevel == ENLogLevel.VERBOSE)
                return;

            log("[ERROR]", method, logText);
        }

        /// <summary>
        /// Attempts to log an Informational message.  The actual logging may or may not occur as a 
        /// funciton of the current Log Level.
        /// NOTE: Please use the logInfo(method, text) method wherever possible
        /// </summary>
        /// <param name="logText">The message to log</param>
        public void logInfo(string logText)
        {
            logInfo("", logText);
        }

        /// <summary>
        /// Attempts to log an Informational message.  The actual logging may or may not occur as a 
        /// funciton of the current Log Level.
        /// </summary>
        /// <param name="method">The location from which the log message was requested</param>
        /// <param name="logText">The message to log</param>
        public void logInfo(string method, string logText)
        {
            if (_logLevel != ENLogLevel.VERBOSE)
                return;

            log("[INFO]", method, logText);
        }

        #endregion
        #region Private Utility functions

        /// <summary>
        /// Main logging method
        /// </summary>
        /// <param name="type">The type of log message: [INFO], [ERROR], [CRITICAL]</param>
        /// <param name="method">The method where logging was requested (allows tracking)</param>
        /// <param name="logText">The text of the log message</param>
        private void log(string type, string method, string logText)
        {
            try
            {
                this.checkLogFileSize();
                _str = File.AppendText(_logFileName);
                _str.WriteLine(now() + ": " + type + method + " : " + logText);
                _str.Close();
            }
            catch (Exception eLog)
            {
                if (eLog.Message.IndexOf("Could not find file") >= 0)
                {
                    try
                    {
                        StreamWriter SW = File.CreateText(_logFileName);
                        SW.WriteLine("Application Log");
                        SW.WriteLine("------------------");
                        SW.Close();
                        this.InitializeLog();
                    }
                    catch
                    {
                        MessageBox.Show("[Logger.Log] Unable to create save log entry:\n" +
                                        eLog.Message + "\n\nLog Text:\n" +
                                        logText);
                    }
                }
            }
        }

        /// <summary>
        /// Provides the date/time string
        /// </summary>
        /// <returns>The current date and time in UTC: "YYYYMMDD.HH:MM:SS"</returns>
        private string now()
        {
            string yr = null;
            string mon = null;
            string day = null;
            string hr = null;
            string min = null;
            string sec = null;

            System.DateTime dtNow = System.DateTime.Now.ToUniversalTime();
            yr = dtNow.Year.ToString();
            mon = dtNow.Month.ToString();
            if (mon.Length < 2)
            {
                mon = "0" + mon;
            }
            day = dtNow.Day.ToString();
            if (day.Length < 2)
            {
                day = "0" + day;
            }
            hr = dtNow.Hour.ToString();
            if (hr.Length < 2)
            {
                hr = "0" + hr;
            }
            min = dtNow.Minute.ToString();
            if (min.Length < 2)
            {
                min = "0" + min;
            }
            sec = dtNow.Second.ToString();
            if (sec.Length < 2)
            {
                sec = "0" + sec;
            }
            string stamp = yr + mon + day + ".";
            stamp = stamp + hr + ":" + min + ":" + sec + " UTC" + "\t";
            return stamp;
        }

        /// <summary>
        /// Creates a string representing a long version of the time
        /// </summary>
        /// <returns>The long version of the time</returns>
        private string nowLong()
        {
            System.DateTime dtNow = System.DateTime.Now.ToUniversalTime();
            return dtNow.ToLongDateString() + " - " + dtNow.ToShortTimeString() + " UTC  ";
        }

        /// <summary>
        /// Ensures that there is sufficient space in the log file
        /// </summary>
        private void checkLogFileSize()
        {
            // Check size of the QWCLog.txt
            FileInfo fi = new FileInfo(_logFileName);
            long currentLogFileSize = fi.Length;
            int temp = getMaxLogFileSize();

            if (temp == 0)
            {
                if (_logLevel != ENLogLevel.NONE)
                {
                    setLogLevel(ENLogLevel.NONE);
                }
                return;
            }
            if (temp < 0)
            {
                if (_logLevel != ENLogLevel.CRITICAL)
                {
                    setLogLevel(ENLogLevel.CRITICAL);
                }
                return;
            }
            if (currentLogFileSize > _maxLogFileSize)
            {
                // delete QWCLogOld.txt if any
                try
                {
                    File.Delete(_oldLogFileName);
                }
                catch
                {
                    //Do nothing
                }

                // rename QWCLog.txt to QWCLogOld.txt
                try
                {
                    File.Move(@_logFileName, _oldLogFileName);
                }
                catch
                {
                    //Do nothing
                }

                // create a new QWCLog.txt
                try
                {
                    this.InitializeLog();
                }
                catch
                {
                    //Do nothing
                }
            }
        }
        private void InitializeLog()
        {
            string dir = Path.GetDirectoryName(_logFileName);

            // Create directory and file, as required...
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            if (File.Exists(_logFileName))
            {
                _str = File.AppendText(_logFileName);
            }
            else
            {
                _str = File.CreateText(_logFileName);
            }

            // Add some initialization string to the log file...
            _str.WriteLine("\r\n\r\nLog file initialized at " + nowLong());
            _str.WriteLine("Timestamp format used: YYYYMMDD.HH:MM:SS UTC");
            string messaging = "";

            // Initialize the logging level and the maximum size...
            _logLevel = getLogLevel();
            _maxLogFileSize = getMaxLogFileSize(true);

            _str.WriteLine("\"" + _productName + "\" version: " + _productVersion + ", has been initialized with its logging " +
                "status to _logLevel = " + _logLevel.ToString() + ".\n" + messaging);
            _str.WriteLine("");
            _str.Close();

            if (_maxLogFileSize <= 1024)
            {
                //If the logging size is pathetically small, it is equivalent to _logLevel = NONE
                if (_logLevel != ENLogLevel.NONE)
                {
                    log("[INFO]", "Logger.checkLogFileSize()", "Log file size only " + _maxLogFileSize + " bytes. Setting LogLevel to NONE for no logging.");
                    setLogLevel(ENLogLevel.NONE);
                }
            }
        }

        /// <summary>
        /// Sets the maximum log file size in the registry
        /// </summary>
        /// <param name="fileSize">value to set</param>
        /// <returns>true if the maximum file size was successful saved in the registry</returns>
        private bool setMaxLogFileSize(int fileSize)
        {
            try
            {
                Microsoft.Win32.RegistryKey rkQBWC = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(Defaults.REG_KEY);
                rkQBWC.SetValue(MaxLogSizeKey, fileSize, RegistryValueKind.DWord);
                rkQBWC.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// Obtains the maximum file size from the registry.  If an entry is not found then the default
        /// value is saved to the registry.   If an error is detected in any of this,  the default
        /// value is returned
        /// </summary>
        /// <param name="bStoreValue">True if the registry value should be forcibly set after being retrieved (handles default case)</param>
        /// <returns>The maximum file size</returns>
        private int getMaxLogFileSize()
        {
            return getMaxLogFileSize(false);
        }

        /// <summary>
        /// Obtains the maximum file size from the registry.  If an entry is not found then the default
        /// value is saved to the registry.   If an error is detected in any of this,  the default
        /// value is returned
        /// </summary>
        /// <param name="bStoreValue">True if the registry value should be forcibly set after being retrieved (handles default case)</param>
        /// <returns>The maximum file size</returns>
        private int getMaxLogFileSize(bool bStoreValue)
        {
            int retValue = Defaults.DEFAULT_LOG_FILESIZE;

            try
            {
                // Attempt to get the MaxFileSize from the registry.  If the key cannot be found there, 
                // create a new key, setting the value to the default value.
                Microsoft.Win32.RegistryKey rkQBWC = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Defaults.REG_KEY, false);
                retValue = (Int32)rkQBWC.GetValue(MaxLogSizeKey, Defaults.DEFAULT_LOG_FILESIZE);
                rkQBWC.Close();

                // Because the default value will be returned if a no registry entry is found (and we'll
                // never know about the lack of an entry), set the value in the registry if so commanded.
                if (bStoreValue)
                    setMaxLogFileSize(retValue);
            }
            catch
            {
                if (bStoreValue)
                    setMaxLogFileSize(retValue);
            }

            return retValue;
        }

        /// <summary>
        /// Set the logging level in the registry
        /// </summary>
        /// <param name="val">The maximum file size</param>
        /// <returns>true if the logging level was saved in the registry</returns>
        private bool setLogLevel(ENLogLevel val)
        {
            try
            {
                Microsoft.Win32.RegistryKey rkQBWC = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(Defaults.REG_KEY);
                rkQBWC.SetValue(LogLevelKey, val.ToString(), RegistryValueKind.String);
                rkQBWC.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Obtains the logging level from the registry.  If an entry is not found then the default
        /// level is saved to the registry.   If an error is detected in any of this,  the default
        /// level is returned
        /// </summary>
        /// <returns>The current logging level</returns>
        private ENLogLevel getLogLevel()
        {
            ENLogLevel retValue = Defaults.DEFAULT_LOG_LEVEL;

            try
            {
                // Attempt to get the LogLevel from the registry.  If the key cannot be found there, 
                // create a new key, setting the value to the default value.
                Microsoft.Win32.RegistryKey rkQBWC = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Defaults.REG_KEY, false);

                if (rkQBWC == null)
                {
                    setLogLevel(retValue);
                    return retValue;
                }

                string retVal = (string)rkQBWC.GetValue(LogLevelKey);
                rkQBWC.Close();

                switch (retVal)
                {
                    case "NONE":
                        retValue = ENLogLevel.NONE;
                        break;
                    case "CRITICAL":
                        retValue = ENLogLevel.CRITICAL;
                        break;
                    case "ERROR":
                        retValue = ENLogLevel.ERROR;
                        break;
                    case "VERBOSE":
                        retValue = ENLogLevel.VERBOSE;
                        break;
                    default:
                        retValue = Defaults.DEFAULT_LOG_LEVEL;
                        setLogLevel(retValue);
                        break;
                }
            }
            catch
            {
                // Don't do anything
            }

            return retValue;
        }
        #endregion
    }
}

