using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;  // For using MessageBox.
using QBSDKEVENTLib; // In order to implement IQBEventCallback.
using System.Runtime.InteropServices;  // For use of the GuidAttribute, ProgIdAttribute and ClassInterfaceAttribute.
using System.Xml; //XML Parsing


namespace SubscribeAndHandleQBEvent
{
	[
      Guid("62447F81-C195-446f-8201-94F0614E49D5"),  // We indicate a specific CLSID for "SubscribeAndHandleQBEvent.EventHandlerObj" for convenience of searching the registry.
      ProgId("SubscribeAndHandleQBEvent.EventHandlerObj"),  // This ProgId is used by default. Not 100% necessary.
      ClassInterface(ClassInterfaceType.None)  
    ]
	public class EventHandlerObj :
        ReferenceCountedObjectBase, // EventHandlerObj is derived from ReferenceCountedObjectBase so that we can track its creation and destruction.
		IQBEventCallback  // this must implement the IQBEventCallback interface.
	{
		
		public EventHandlerObj()
		{
			// ReferenceCountedObjectBase constructor will be invoked.
            Console.WriteLine("EventHandlerObj constructor.");
		}

        ~EventHandlerObj()
		{
			// ReferenceCountedObjectBase destructor will be invoked.
            Console.WriteLine("EventHandlerObj destructor.");
		}

        //Call back function which would be invoked from the QB
		public void inform(string strMessage)
		{
            try
            {
                StringBuilder sb = new StringBuilder(strMessage);
                XmlDocument outputXMLDoc = new XmlDocument();
                outputXMLDoc.LoadXml(strMessage);
                XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("QBXMLEvents");
                XmlNode childNode = qbXMLMsgsRsNodeList.Item(0).FirstChild;
                
                // handle the event based on type of event
                switch (childNode.Name)
                {
                    case "DataEvent":
                        //Handle Data Event Here
                        MessageBox.Show(sb.ToString(), "DATA EVENT - From QB");
                        break;

                    case "UIExtensionEvent":
                        //Handle UI Extension Event HERE
                        MessageBox.Show(sb.ToString(), "UI EXTENSION EVENT - From QB");
                        break;

                    default:
                        MessageBox.Show(sb.ToString(), "Response From QB");
                        break;
                }
                               
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error in processing the response from QB - " + ex.Message);
            }
		}
	}

    class EventHandlerObjClassFactory : ClassFactoryBase
	{
		public override void virtual_CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
		{
            Console.WriteLine("EventHandlerObjClassFactory.CreateInstance().");
			Console.WriteLine("Requesting Interface : " + riid.ToString());

            if (riid == Marshal.GenerateGuidForType(typeof(IQBEventCallback)) ||
                riid == SubscribeAndHandleQBEvent.IID_IDispatch ||
                riid == SubscribeAndHandleQBEvent.IID_IUnknown)
			{
                EventHandlerObj EventHandlerObj_New = new EventHandlerObj();

                ppvObject = Marshal.GetComInterfaceForObject(EventHandlerObj_New, typeof(IQBEventCallback));
			} 
			else
			{
				throw new COMException("No interface",  unchecked((int) 0x80004002));
			}
		}
	}
}