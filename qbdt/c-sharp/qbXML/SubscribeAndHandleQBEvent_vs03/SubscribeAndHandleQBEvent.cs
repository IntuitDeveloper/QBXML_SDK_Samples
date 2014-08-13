using System;
using System.Drawing;
using System.Windows.Forms;
using System.Xml;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Text;
using System.Threading;
using Interop.QBXMLRP2;

namespace SubscribeAndHandleQBEvent
{
     [Flags]
     enum COINIT : uint
     {
         /// Initializes the thread for multi-threaded object concurrency.
         COINIT_MULTITHREADED = 0x0,
         /// Initializes the thread for apartment-threaded object concurrency. 
         COINIT_APARTMENTTHREADED = 0x2,
         /// Disables DDE for Ole1 support.
         COINIT_DISABLE_OLE1DDE = 0x4,
         /// Trades memory for speed.
         COINIT_SPEED_OVER_MEMORY = 0x8
     }

     [Flags]
     enum CLSCTX : uint
     {
         CLSCTX_INPROC_SERVER = 0x1,
         CLSCTX_INPROC_HANDLER = 0x2,
         CLSCTX_LOCAL_SERVER = 0x4,
         CLSCTX_INPROC_SERVER16 = 0x8,
         CLSCTX_REMOTE_SERVER = 0x10,
         CLSCTX_INPROC_HANDLER16 = 0x20,
         CLSCTX_RESERVED1 = 0x40,
         CLSCTX_RESERVED2 = 0x80,
         CLSCTX_RESERVED3 = 0x100,
         CLSCTX_RESERVED4 = 0x200,
         CLSCTX_NO_CODE_DOWNLOAD = 0x400,
         CLSCTX_RESERVED5 = 0x800,
         CLSCTX_NO_CUSTOM_MARSHAL = 0x1000,
         CLSCTX_ENABLE_CODE_DOWNLOAD = 0x2000,
         CLSCTX_NO_FAILURE_LOG = 0x4000,
         CLSCTX_DISABLE_AAA = 0x8000,
         CLSCTX_ENABLE_AAA = 0x10000,
         CLSCTX_FROM_DEFAULT_CONTEXT = 0x20000,
         CLSCTX_INPROC = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
         CLSCTX_SERVER = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER,
         CLSCTX_ALL = CLSCTX_SERVER | CLSCTX_INPROC_HANDLER
     }

     [Flags]
     enum REGCLS : uint
     {
         REGCLS_SINGLEUSE = 0,
         REGCLS_MULTIPLEUSE = 1,
         REGCLS_MULTI_SEPARATE = 2,
         REGCLS_SUSPENDED = 4,
         REGCLS_SURROGATE = 8
     }

     // We import the POINT structure because it is referenced
     // by the MSG structure.
     [ComVisible(false)]
     [StructLayout(LayoutKind.Sequential)]
     public struct POINT
     {
         public int X;
         public int Y;

         public POINT(int x, int y)
         {
             this.X = x;
             this.Y = y;
         }

         public static implicit operator Point(POINT p)
         {
             return new Point(p.X, p.Y);
         }

         public static implicit operator POINT(Point p)
         {
             return new POINT(p.X, p.Y);
         }
     }

    // We import the MSG structure because it is referenced 
    // by the GetMessage(), TranslateMessage() and DispatchMessage()
    // Win32 APIs.
    [ComVisible(false)]
    [StructLayout(LayoutKind.Sequential)]
    public struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
    }

    // Note that SubscribeAndHandleQBEvent is NOT declared as public.
    // This is so that it will not be exposed to COM when we call regasm
    // or tlbexp.
    class SubscribeAndHandleQBEvent
    {

        enum QBSubscriptionType { Data, UI, UIExtension };
        static string strAppName = "QB C# Event Sample";

        // CoInitializeEx() can be used to set the apartment model
        // of individual threads.
        [DllImport("ole32.dll")]
        static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

        // CoUninitialize() is used to uninitialize a COM thread.
        [DllImport("ole32.dll")]
        static extern void CoUninitialize();

        // PostThreadMessage() allows us to post a Windows Message to
        // a specific thread (identified by its thread id).
        // We will need this API to post a WM_QUIT message to the main 
        // thread in order to terminate this application.
        [DllImport("user32.dll")]
        static extern bool PostThreadMessage(uint idThread, uint Msg, UIntPtr wParam,
            IntPtr lParam);

        // GetCurrentThreadId() allows us to obtain the thread id of the
        // calling thread. This allows us to post the WM_QUIT message to
        // the main thread.
        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        // We will be manually performing a Message Loop within the main thread
        // of this application. Hence we will need to import GetMessage(), 
        // TranslateMessage() and DispatchMessage().
        [DllImport("user32.dll")]
        static extern bool GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin,
            uint wMsgFilterMax);

        [DllImport("user32.dll")]
        static extern bool TranslateMessage([In] ref MSG lpMsg);

        [DllImport("user32.dll")]
        static extern IntPtr DispatchMessage([In] ref MSG lpmsg);

        // Define two common GUID objects for public usage.
        public static Guid IID_IUnknown = new Guid("{00000000-0000-0000-C000-000000000046}");
        public static Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

        protected static uint m_uiMainThreadId;  // Stores the main thread's thread id.
        protected static int m_iObjsInUse;  // Keeps a count on the total number of objects alive.
        protected static int m_iServerLocks;// Keeps a lock count on this application.

        // This property returns the main thread's id.
        public static uint MainThreadId
        {
            get
            {
                return m_uiMainThreadId;
            }
        }

        // This method performs a thread-safe incrementation of the objects count.
        public static int InterlockedIncrementObjectsCount()
        {
            Console.WriteLine("InterlockedIncrementObjectsCount()");
            // Increment the global count of objects.
            return Interlocked.Increment(ref m_iObjsInUse);
        }

        // This method performs a thread-safe decrementation the objects count.
        public static int InterlockedDecrementObjectsCount()
        {
            Console.WriteLine("InterlockedDecrementObjectsCount()");
            // Decrement the global count of objects.
            return Interlocked.Decrement(ref m_iObjsInUse);
        }

        // Returns the total number of objects alive currently.
        public static int ObjectsCount
        {
            get
            {
                lock (typeof(SubscribeAndHandleQBEvent))
                {
                    return m_iObjsInUse;
                }
            }
        }

        // This method performs a thread-safe incrementation the 
        // server lock count.
        public static int InterlockedIncrementServerLockCount()
        {
            Console.WriteLine("InterlockedIncrementServerLockCount()");
            // Increment the global lock count of this server.
            return Interlocked.Increment(ref m_iServerLocks);
        }

        // This method performs a thread-safe decrementation the 
        // server lock count.
        public static int InterlockedDecrementServerLockCount()
        {
            Console.WriteLine("InterlockedDecrementServerLockCount()");
            // Decrement the global lock count of this server.
            return Interlocked.Decrement(ref m_iServerLocks);
        }

        // Returns the current server lock count.
        public static int ServerLockCount
        {
            get
            {
                lock (typeof(SubscribeAndHandleQBEvent))
                {
                    return m_iServerLocks;
                }
            }
        }

        // AttemptToTerminateServer() will check to see if 
        // the objects count and the server lock count has
        // both dropped to zero.
        // If so, we post a WM_QUIT message to the main thread's
        // message loop. This will cause the message loop to
        // exit and hence the termination of this application.
        public static void AttemptToTerminateServer()
        {
            lock (typeof(SubscribeAndHandleQBEvent))
            {
                Console.WriteLine("AttemptToTerminateServer()");

                // Get the most up-to-date values of these critical data.
                int iObjsInUse = ObjectsCount;
                int iServerLocks = ServerLockCount;

                // Print out these info for debug purposes.
                StringBuilder sb = new StringBuilder("");
                sb.AppendFormat("m_iObjsInUse : {0}. m_iServerLocks : {1}", iObjsInUse, iServerLocks);
                Console.WriteLine(sb.ToString());

                if ((iObjsInUse > 0) || (iServerLocks > 0))
                {
                    Console.WriteLine("There are still referenced objects or the server lock count is non-zero.");
                }
                else
                {
                    UIntPtr wParam = new UIntPtr(0);
                    IntPtr lParam = new IntPtr(0);
                    Console.WriteLine("PostThreadMessage(WM_QUIT)");
                    PostThreadMessage(MainThreadId, 0x0012, wParam, lParam);
                }
            }
        }

        // ProcessArguments() will process the command-line arguments
        // of this application. 
        // If the return value is true, we carry
        // on and start this application.
        // If the return value is false, we terminate
        // this application immediately.
        protected static bool ProcessArguments(string[] args)
        {
            bool bRet = true;

            if (args.Length > 0)
            {
                RegistryKey key = null;
                RegistryKey key2 = null;

                switch (args[0].ToLower())
                {
                    case "-h": //Help
                    case "/h": //Help
                        DisplayUsage();
                        bRet = false;
                        break;

                    case "-embedding": //COM sends this as the argument when starting a out of proc server
                        Console.WriteLine("Request to start as out-of-process COM server.");
                        break;
                        
                    case "-regserver":
                    case "/regserver":
                        try
                        {
                            key = Registry.ClassesRoot.CreateSubKey("CLSID\\" + Marshal.GenerateGuidForType(typeof(EventHandlerObj)).ToString("B"));
                            key2 = key.CreateSubKey("LocalServer32");
                            key2.SetValue(null, Application.ExecutablePath);
                            
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error while registering the server:\n" + ex.ToString());
                        }
                        finally
                        {
                            if (key != null)
                                key.Close();
                            if (key2 != null)
                                key2.Close();
                        }
                        bRet = false;
                        break;

                    case "-unregserver":
                    case "/unregserver":
                        try
                        {
                            key = Registry.ClassesRoot.OpenSubKey("CLSID\\" + Marshal.GenerateGuidForType(typeof(EventHandlerObj)).ToString("B"), true);
                            key.DeleteSubKey("LocalServer32");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error while unregistering the server:\n" + ex.ToString());
                        }
                        finally
                        {
                            if (key != null)
                                key.Close();
                            if (key2 != null)
                                key2.Close();
                            //unsubscribe from QB events
                            UnsubscribeForEvents(QBSubscriptionType.Data, true); // don't display error 
                            UnsubscribeForEvents(QBSubscriptionType.UIExtension, true); // don't display error 
                        }
                        bRet = false;
                        break;

                    //subscribing for data Event
                    case "-d":
                    case "/d":
                         //Subscribe for Quick Books Data Event - Customer add/Modify/Delete event, if not already subscribed
                         SubscribeForEvents(QBSubscriptionType.Data,String.Empty);
                         bRet = false;
                         break;

                     //subscribing for UI Extension Event
                     case "-u":
                     case "/u":
                         {
                             //Subscribe for UI Extension event - Adding a menu item under Customer Menu in QB.
                             //Get the Menu Item Name from Arguments
                             if (args.Length < 2)
                             {
                                 // Menu item name is not provided.
                                 // Display Usage and exit
                                 DisplayUsage();

                             }
                             else
                             {
                                 SubscribeForEvents(QBSubscriptionType.UIExtension,args[1]);
                             }
                             bRet = false;
                             break;
                         }

                     case "-dd":
                     case "/dd":
                         //unsubscribe for Quick Books Data Event
                         UnsubscribeForEvents(QBSubscriptionType.Data,false);
                         bRet = false;
                         break;

                     case "-ud":
                     case "/ud":
                         //unsubscribe for Quick Books UIExtension Event
                         UnsubscribeForEvents(QBSubscriptionType.UIExtension,false);
                         bRet = false;
                         break;

                     default:
                         DisplayUsage();
                         bRet = false;
                        break;
                }
            }

            return bRet;
        }

        private static void DisplayUsage()
        {
            string strUsage = "Usage";
            strUsage += "\n -regserver \n\t:register as COM out of proc server";
            strUsage += "\n -unregserver \n\t: unregister COM out of proc server. Also unsubscribes from all the events.";
            strUsage += "\n -d		\n\t: subscribe for customer add/modify/delete data event";
            strUsage += "\n -u <Menu Name>   \n\t: subscribe for UI extnesion event. <Menu Name> will appear under customers menu in QB";
            strUsage += "\n -dd 		\n\t: unsubscribe for customer data event";
            strUsage += "\n -ud 		\n\t: unsubscribe for UI extension event";
            strUsage += "\n";

            Console.Write(strUsage);
	     }
        //Subscribes this application to listen for Data event or UI extension event
        private static void SubscribeForEvents(QBSubscriptionType strType, string strData)
        {
            RequestProcessor2Class qbRequestProcessor;
            try
            {
                // Get an instance of the qbXMLRP Request Processor and
                // call OpenConnection if that has not been done already.
                qbRequestProcessor = new RequestProcessor2Class();
                qbRequestProcessor.OpenConnection("", strAppName);

                StringBuilder strRequest = new StringBuilder();
                switch (strType)
                {
                    case QBSubscriptionType.Data:
                        strRequest = new StringBuilder(GetDataEventSubscriptionAddXML());
                        break;

                    case QBSubscriptionType.UIExtension:
                        strRequest = new StringBuilder(GetUIExtensionSubscriptionAddXML(strData));
                        break;

                    default:
                        return;
                }
                                
                string strResponse = qbRequestProcessor.ProcessSubscription(strRequest.ToString());
                
                //Parse the XML response to check the status
                XmlDocument outputXMLDoc = new XmlDocument();
                outputXMLDoc.LoadXml(strResponse);
                XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("DataEventSubscriptionAddRs");
                if (qbXMLMsgsRsNodeList.Count == 1)
                {
                     XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
                    //get the status Code, info and Severity
                    string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value;
                    string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                    string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;

                    if ((retStatusCode != "0") && (retStatusCode != "3180"))// 3180 : if subscription already subscribed. NOT A NEAT WAY TO DO THIS, NEED TO EXPLORE THIS
                    {
                        Console.WriteLine("Error while subscribing for events\n\terror Code - {0},\n\tSeverity - {1},\n\tError Message - {2}\n", retStatusCode, retStatusSeverity, retStatusMessage);
                        return;
                    }
                } 
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while registering for QB events - " + ex.Message);
                qbRequestProcessor = null;
                return ;
            }
        }

        //Unsubscribes this application from listening to add/modify/delete custmor event
        private static void UnsubscribeForEvents(QBSubscriptionType strType,bool bSilent)
        {
            RequestProcessor2Class qbRequestProcessor;
            try
            {
                // Get an instance of the qbXMLRP Request Processor and
                // call OpenConnection if that has not been done already.
                qbRequestProcessor = new RequestProcessor2Class();
                qbRequestProcessor.OpenConnection("", strAppName);

                StringBuilder strRequest = new StringBuilder(GetSubscriptionDeleteXML(strType));
                string strResponse = qbRequestProcessor.ProcessSubscription(strRequest.ToString());

                //Parse the XML response to check the status
                XmlDocument outputXMLDoc = new XmlDocument();
                outputXMLDoc.LoadXml(strResponse);
                XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("SubscriptionDelRs");
                       
                XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
                //get the status Code, info and Severity
                string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;

                if ((retStatusCode != "0")&&(!bSilent))
                {
                    Console.WriteLine("Error while unsubscribing from events\n\terror Code - {0},\n\tSeverity - {1},\n\tError Message - {2}\n", retStatusCode, retStatusSeverity, retStatusMessage);
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while unsubscribing from QB events - " + ex.Message);
                qbRequestProcessor = null;
                return ;
            }
            return;
        }
        

        // This Method return the qbXML for the Subscribing this application to QB for listening 
        // to customer add/modify/delete event.
        private static string GetDataEventSubscriptionAddXML()
        {
            //Create the qbXML request
            XmlDocument requestXMLDoc = new XmlDocument();
            requestXMLDoc.AppendChild(requestXMLDoc.CreateXmlDeclaration("1.0", null, null));
            requestXMLDoc.AppendChild(requestXMLDoc.CreateProcessingInstruction("qbxml", "version=\"5.0\""));
            XmlElement qbXML = requestXMLDoc.CreateElement("QBXML");
            requestXMLDoc.AppendChild(qbXML);

            //subscription Message request
            XmlElement qbXMLMsgsRq = requestXMLDoc.CreateElement("QBXMLSubscriptionMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
  
            //Data Event Subscription ADD request
            XmlElement dataEventSubscriptionAddRq = requestXMLDoc.CreateElement("DataEventSubscriptionAddRq");
            qbXMLMsgsRq.AppendChild(dataEventSubscriptionAddRq);


            //Data Event Subscription ADD
            XmlElement dataEventSubscriptionAdd = requestXMLDoc.CreateElement("DataEventSubscriptionAdd");
            dataEventSubscriptionAddRq.AppendChild(dataEventSubscriptionAdd);

            //Add Subscription ID
            dataEventSubscriptionAdd.AppendChild(requestXMLDoc.CreateElement("SubscriberID")).InnerText = "{8327c7fc-7f05-41ed-a5b4-b6618bb27bf1}";

            //Add COM CallbackInfo
            XmlElement comCallbackInfo = requestXMLDoc.CreateElement("COMCallbackInfo");
            dataEventSubscriptionAdd.AppendChild(comCallbackInfo);

            //Appname and CLSID
            comCallbackInfo.AppendChild(requestXMLDoc.CreateElement("AppName")).InnerText = strAppName;
            comCallbackInfo.AppendChild(requestXMLDoc.CreateElement("CLSID")).InnerText = "{62447F81-C195-446f-8201-94F0614E49D5}";

            //Delivery Policy
            dataEventSubscriptionAdd.AppendChild(requestXMLDoc.CreateElement("DeliveryPolicy")).InnerText = "DeliverAlways";

            //track lost events
            dataEventSubscriptionAdd.AppendChild(requestXMLDoc.CreateElement("TrackLostEvents")).InnerText = "All";


            //  ListEventSubscription
            XmlElement listEventSubscription = requestXMLDoc.CreateElement("ListEventSubscription");
            dataEventSubscriptionAdd.AppendChild(listEventSubscription);

            //Add Customer List and operations
            listEventSubscription.AppendChild(requestXMLDoc.CreateElement("ListEventType")).InnerText = "Customer";
            listEventSubscription.AppendChild(requestXMLDoc.CreateElement("ListEventOperation")).InnerText = "Add";
            listEventSubscription.AppendChild(requestXMLDoc.CreateElement("ListEventOperation")).InnerText = "Modify";
            listEventSubscription.AppendChild(requestXMLDoc.CreateElement("ListEventOperation")).InnerText = "Delete";

            string strRetString = requestXMLDoc.OuterXml;
            LogXmlData(@"C:\Temp\DataEvent.xml", strRetString);
            return strRetString;
            
        }



        // This Method return the qbXML for the Adding a UI extension to the customer menu.
        // Event will be received any time the menu is clicked 
        private static string GetUIExtensionSubscriptionAddXML(string strMenuName)
        {
            //Create the qbXML request
            XmlDocument requestXMLDoc = new XmlDocument();
            requestXMLDoc.AppendChild(requestXMLDoc.CreateXmlDeclaration("1.0", null, null));
            requestXMLDoc.AppendChild(requestXMLDoc.CreateProcessingInstruction("qbxml", "version=\"5.0\""));
            XmlElement qbXML = requestXMLDoc.CreateElement("QBXML");
            requestXMLDoc.AppendChild(qbXML);

            //subscription Message request
            XmlElement qbXMLMsgsRq = requestXMLDoc.CreateElement("QBXMLSubscriptionMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);

            //UI Extension Subscription ADD request
            XmlElement uiExtSubscriptionAddRq = requestXMLDoc.CreateElement("UIExtensionSubscriptionAddRq");
            qbXMLMsgsRq.AppendChild(uiExtSubscriptionAddRq);


            //UI Extension Subscription ADD
            XmlElement uiExtEventSubscriptionAdd = requestXMLDoc.CreateElement("UIExtensionSubscriptionAdd");
            uiExtSubscriptionAddRq.AppendChild(uiExtEventSubscriptionAdd);

            //Add Subscription ID
            uiExtEventSubscriptionAdd.AppendChild(requestXMLDoc.CreateElement("SubscriberID")).InnerText = "{8327c7fc-7f05-41ed-a5b4-b6618bb27bf1}";

            //Add COM CallbackInfo
            XmlElement comCallbackInfo = requestXMLDoc.CreateElement("COMCallbackInfo");
            uiExtEventSubscriptionAdd.AppendChild(comCallbackInfo);

            //Appname and CLSID
            comCallbackInfo.AppendChild(requestXMLDoc.CreateElement("AppName")).InnerText = strAppName;
            comCallbackInfo.AppendChild(requestXMLDoc.CreateElement("CLSID")).InnerText = "{62447F81-C195-446f-8201-94F0614E49D5}";

            
            //  MenuEventSubscription
            XmlElement menuExtensionSubscription = requestXMLDoc.CreateElement("MenuExtensionSubscription");
            uiExtEventSubscriptionAdd.AppendChild(menuExtensionSubscription);

            //Add To menu Item // To Cusomter Menu
            menuExtensionSubscription.AppendChild(requestXMLDoc.CreateElement("AddToMenu")).InnerText = "Customers";


            XmlElement menuItem= requestXMLDoc.CreateElement("MenuItem");
            menuExtensionSubscription.AppendChild(menuItem);

            //Add Menu Name
            menuItem.AppendChild(requestXMLDoc.CreateElement("MenuText")).InnerText = strMenuName;
            menuItem.AppendChild(requestXMLDoc.CreateElement("EventTag")).InnerText = "menu_"+ strMenuName;


            XmlElement displayCondition = requestXMLDoc.CreateElement("DisplayCondition");
            menuItem.AppendChild(displayCondition);

            displayCondition.AppendChild(requestXMLDoc.CreateElement("VisibleIf")).InnerText = "HasCustomers";
            displayCondition.AppendChild(requestXMLDoc.CreateElement("EnabledIf")).InnerText = "HasCustomers";
          

            string strRetString = requestXMLDoc.OuterXml;
            LogXmlData(@"C:\Temp\UIExtension.xml", strRetString);
            return strRetString;

        }

        // This Method return the qbXML for the deleting the event subscribing for this application
        // from QB
        private static string GetSubscriptionDeleteXML(QBSubscriptionType subscriptionType)
        {
            //Create the qbXML request
            XmlDocument requestXMLDoc = new XmlDocument();
            requestXMLDoc.AppendChild(requestXMLDoc.CreateXmlDeclaration("1.0", null, null));
            requestXMLDoc.AppendChild(requestXMLDoc.CreateProcessingInstruction("qbxml", "version=\"5.0\""));
            XmlElement qbXML = requestXMLDoc.CreateElement("QBXML");
            requestXMLDoc.AppendChild(qbXML);

            //subscription Message request
            XmlElement qbXMLMsgsRq = requestXMLDoc.CreateElement("QBXMLSubscriptionMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);

            //Data Event Subscription ADD request
            XmlElement qbSubscriptionDelRq = requestXMLDoc.CreateElement("SubscriptionDelRq");
            qbXMLMsgsRq.AppendChild(qbSubscriptionDelRq);

            //Subscription ID
            qbSubscriptionDelRq.AppendChild(requestXMLDoc.CreateElement("SubscriberID")).InnerText = "{8327c7fc-7f05-41ed-a5b4-b6618bb27bf1}";

            //Subscription Type
            qbSubscriptionDelRq.AppendChild(requestXMLDoc.CreateElement("SubscriptionType")).InnerText = subscriptionType.ToString();


            string strRetString = requestXMLDoc.OuterXml;
            LogXmlData(@"C:\Temp\Unsubscribe.xml", strRetString);
            return strRetString;

        }


        // Used only for debug purpose
        // strFile Name should have complete Path of the file too
        private static void LogXmlData(string strFile,string strXML)
        {
            System.IO.StreamWriter sw = new System.IO.StreamWriter(strFile);
            sw.WriteLine(strXML);
            sw.Flush();
            sw.Close();
         }


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if (!ProcessArguments(args))
                {
                    return;
                }

                // Initialize critical member variables.
                m_iObjsInUse = 0;
                m_iServerLocks = 0;
                m_uiMainThreadId = GetCurrentThreadId();

                // Register the EventHandlerObjClassFactory.
                EventHandlerObjClassFactory factory = new EventHandlerObjClassFactory();
                factory.ClassContext = (uint)CLSCTX.CLSCTX_LOCAL_SERVER;
                factory.ClassId = Marshal.GenerateGuidForType(typeof(EventHandlerObj));
                factory.Flags = (uint)REGCLS.REGCLS_MULTIPLEUSE | (uint)REGCLS.REGCLS_SUSPENDED;
                factory.RegisterClassObject();
                ClassFactoryBase.ResumeClassObjects();

                Console.WriteLine("Waiting for QB Customer Add Event .....\n");
                // Start the message loop.
                MSG msg;
                IntPtr null_hwnd = new IntPtr(0);
                while (GetMessage(out msg, null_hwnd, 0, 0) != false)
                {
                    TranslateMessage(ref msg);
                    DispatchMessage(ref msg);
                }
                Console.WriteLine("Out of message loop.");

                // Revoke the class factory immediately.
                // Don't wait until the thread has stopped before
                // we perform revokation.
                factory.RevokeClassObject();
                Console.WriteLine("EventHandlerObjClassFactory Revoked.");

                // Just an indication that this COM EXE Server is stopped.
                Console.WriteLine("Press [ENTER] to exit.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error in program - " + ex.Message);
            }
        }
            
    }
}
