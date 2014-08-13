The invoice sample is meant to be run against QuickBooks 2003 Canadian version.  It expects
QuickBooks to be running with a company file open and the Admin user logged in. The user 
can send an invoice to QuickBooks containing up to 3 items.  The item types used
by this sample app are Item Service, Item Non Inventory and Item Other Charge. This app 
will work on a company file that has multicurrency turned on or off.  

This show how to retrieve the qbXML versions supported by QuickBooks.  It verifies if QuickBooks
is compatible with the sample app.  You could extend this to do an app that works with both
Canadian and US version of QuickBooks (not shown in this example).  The check is done in the 
OpenConnection function in the qbooks.bas module. 