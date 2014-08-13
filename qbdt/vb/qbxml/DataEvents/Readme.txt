This application demonstrates the use of DataEvents to keep a simple "database"
(implemented as a Windows ListView common control) of customers in sync with
QuickBooks' customer list using DataEvents, and shows the <b><em>careful</em><b>
use of Data Extensions as a direct result of events received.  Please see the
code commentary for details regarding the discretion and care required if you
plan to do anything other than Query data as a direct result of an event.

This application is made up of 3 executables:

QBDataEventSubscriber -- Subscribes the application to receive data events
                         from QuickBooks.  This is the sort of application you
                         would build as a Custom Action for an MSI installer to
                         subscribe your application to events when you 
			 install your app and unsubscribe when your app 
			 is uninstalled.

			 You MUST run this executable to subscribe the 
			 EventManager to events for the others to do 
			 anything interesting.

QBDataEventManager    -- Handles events from QuickBooks and supplies them 
                         to the main application via a second COM method 
                         -- also has a UI to display the incoming event XML.

			 This executable will be started automatically by
			 QuickBooks (or the DataEventApp) as it is a COM
			 server.

			 NOTE: This executable must be registered as a COM
			 server prior to use and before you can successfully
			 run the QBDataEventSubscriber.  Simply run 
			 QBDataEventManager.exe /REGSERVER.

QBDataEventApp        -- Provides the UI for this application, all this
                         application does is display a form with a list 
   	                 view of the customer list from QuickBooks, as 
		         customers are changed, added, and deleted in 
		         QuickBooks (or by other applications) those 
		         changes are reflected in the list view.  Changed
		         customers turn red, new customers are blue, 
		         deleted customers are yellow and the name is 
		         replaced by #deleted.

		         The first time you run this application it will
		         take a little while as it will create a data 
			 extension and each customer record will be 
			 extended to store the position the customer
			 record holds in the ListView, from then on that
			 position belongs to that customer!  This means that
			 as customers are added and deleted in QuickBooks
			 records will not be in alphabetical order in the
			 listview and records deleted since the first time
			 the application ran will be shown in the list 
			 view as #deleted. 


In short, to successfully run this sample:

 a. Register the event manager with COM
 b. Run the DataEventSubscriber to subscribe to events
 c. Restart QB company file, so that the subscription is acknowledged
 d. Start QBDataEventApp.exe
 e. Modify/add/delete customers in QB and notice that they get updated 
    in QBDataEventApp as well.


