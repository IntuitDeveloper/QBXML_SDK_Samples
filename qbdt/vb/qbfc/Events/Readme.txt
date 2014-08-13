The QBFCEventsSubscriber and QBFCEventsCallback samples work together to show how to implement data and UI event subscriptions and callback handling using the QBFC library. 


The QBFCEventsSubscriber application subscribes to the CustomerAdd and company file Close events. 
    Note: In order to subscribe, the QBFCEventsCallback application must first be registered.  You must run this executable to subscribe to Events for the QBFCEventsCallback to do anything interesting.


The QBFCEventsCallback application parses and processes the events. For the CustomerAdd event, it simply displays the customer that was added.  It uses the ListID returned in the CustomerAdd event to query QuickBooks for that customer's FullName and then displays a non blocking UI message to the user about the added customer. (Blocking UI messages can prevent the reception and handling of subsequent events, including company Close events.) 

The QBFCEventsCallback also contains a "Query" to get all the customers from the QuickBooks company file.  This functionality is available until a company file close event is received.  When a company file Close event is received, the "Query" button is greyed out and the user can no longer query QuickBooks.

    NOTE: This executable must be registered as a COM
    server prior to use and before you can successfully
    subscribe to or receive events.  Simply run QBFCEventsCallback.exe /REGSERVER.
