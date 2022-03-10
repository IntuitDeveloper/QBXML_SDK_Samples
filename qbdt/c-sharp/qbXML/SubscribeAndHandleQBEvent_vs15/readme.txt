The purpose of this sample is to explain the steps involved in creating a 
C# out-of-proc server and registering it with QB to listen for events. QB SDK requires COM components to 
be implemented as a out of proc server in order to listen for QB events but in C# it is normall only possible 
to register COM component as in-proc servers. In order to create a COM-like out of proc Server in C#, we need 
to do some things which usually get done behind the scene in ATL/COM. 

The main purpose of this sample is to just show those steps.

PLEASE NOTE THAT SOME OF THIS CODE HAS BEEN TAKEN FROM CODE PROJECT 
(http://www.codeproject.com/useritems/BuildCOMServersInDotNet.asp) and has been modified for this sample.

This is not a generic sample of creating COM out of proc servers in C# as it doesn't actually 
follow the complete rules of COM, but it is specific to the event subscription in QB.

This sample registers a C# COM object as out of proc server. This program can be subscribed to 
QB data and UI extension events. It can listen for Customer Add, delete and modify event. It can also 
add a new menu item under the 'customers' menu in QB and listen for click event of this menu item. 
If registered for these events, this exe is launched (if not already running) on any of the registered 
event and a message box is displayed which shows the qbXML response for the event.



Usage:
---------
SubscribeAndHandleQBEvent.exe  -regserver
	This will register the COM out of proc server. This SHOULD always be the first step to use this sample. 
	This doesn't keep the exe running and comes out once the registration is done.

SubscribeAndHandleQBEvent.exe  -unregserver
	This un-registers the COM server and also unsubscribe from QB events ( if any).

SubscribeAndHandleQBEvent.exe  -d
	Registers for QB data event for customer add, modify and delete

SubscribeAndHandleQBEvent.exe  -u <Menu item Name>
	Registers for UI Extension events. Adds a new menu item <Menu item Name> under the customer menu in QB 
	and registers for this menu item events.

SubscribeAndHandleQBEvent.exe  -dd
	Un-registers from QB data event for customer.

SubscribeAndHandleQBEvent.exe  -ud
	Un-registers from QB UI extension events.

SubscribeAndHandleQBEvent.exe
	Runs the exe and this exe will keep running (and waiting for QB events, if it has subscribed 
	for any events) unless user stops it.

Running the sample:
------------------------
Please install latest QBSDK.

Before running SubscribeAndHandleQBEvent.exe, this exe need to be registered as COM out of proc server. 
Run SubscribeAndHandleQBEvent.exe -h to see the various options with the exe.
Also make sure that .NET runtime is installed on the machine, and QuickBooks need to be restarted every 
time you add a new menu item event subscription.

Follow these steps to run the sample.
	- Copy the exe to the QB install directory.
	- Register the EXE as out of proc server [See usage above]
	- Subscribe for Data and/or UI extension events.  [See usage above]
	- Start QB and accept the SDK request. 
	- Modify or add customer or click on newly added menu item
	- Message box should appear with the event qbxml


Files and Classes:
---------------------------
IClassFactory.cs
	Defines the IClassFactory interface, (This is same as defined by COM library and have been given the same GUID)
ClassFactoryBase.cs
	This class gives a default implementation of for the Class factory and also defines methods to register and un-register class factories when the exe starts.

ReferenceCountedObject.cs
	This class keeps the reference count of the COM object. Any COM object class in this server inherits from this.

SubscribeAndHandleQBEvent.cs
	This is the main class, which handles registering/un-registering of exe as an out-of-proc server. This also keeps the reference count on the server and object count so that server can be unloaded if there are  
no reference to objects or factories. Also on startup it starts a message pump to keep the exe running.

EventHandlerObj.cs
	This is the COM object which is invoked by QB for processing the Customer event. This class implements the IQBEventCallback interface.


Modification:
---------------------
For Event Processing:
	- To process the data event or UI menu click event, modify the inform(string strMessage) method in EventHandlerObj.cs. Make changes in switch statement to add your code for processing certain type of event.

For Data Event subscription:
	- To change the code to subscribe to events other than customer events, modify GetDataEventSubscriptionAddXML function in SubscribeAndHandleQBEvents.cs. The xml can be modified to subscribe to any other list (other then customer) event and any action on the list.

For UI Extension subscription:
	- To add the menu item under any other QB menu modify the xml in GetUIExtensionSubscriptionAddXML function in SubscribeAndHandleQBEvents.cs. Refer to QB XML on screen reference for more details on qbXML.




