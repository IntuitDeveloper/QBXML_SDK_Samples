This sample is a C# ASP.NET web service application that communicates 
with QuickBooks via QBWebConnector. The sample focuses primarily on 
demonstrating how to setup all web service web methods to run against 
QBWebConnector and does not focus on any particular use case. 

For simplicity, it sends three request XMLs: 
CustomerQuery, InvoiceQuery and BillQuery. 


Running the sample
------------------
You will need to rebuild the sample with your own user name and password,
as described in the comments at the beginning of the source code for this program!

You will need Microsoft Internet Information Services (IIS) installed 
on your system to be able to run this sample.

This sample assumes that you have configured IIS with ASP.NET and 
have a functional system to deploy this web service sample. If you have 
not yet configured ASP.NET with IIS, you may need to run the 
following command from c:\windows\Microsoft.NET\Framework\
your_asp_dot_net_version path: aspnet_regiis /i 

This will help avoid the occasional message from microsoft development 
environment such as "VS.NET has detected that the specified web server 
is not running ASP.NET version 1.1. You will be unable to run ASP.NET 
web applications or services). 



Copy the entire directory to inetpub\wwwroot path.  Before running 
this web service sample against QBWebConnector make sure that IIS and 
.NET runtime is installed on the machine, and QuickBooks is running 
with a company file opened. 

As with any web service application that writes to the directory, 
configure permission for ASP.NET user account so it can read/write 
in the directory. Here is one way to set permission for ASP.NET user 
in Windows XP: -

Start - ControlPanel - AdministrativeTools - ComputerManagement - 
LocalUsersAndGroups - Users - RightClick on ASPNET - Properties - 
MemberOf - Administrators group (to be able to write to files).

From within IISConsole,
- WCWebService -- Properties 
- Read, Write, LogVisits, IndexThisResource -- Checked
- In Directory Security Tab -- Annonymous Access + Integrated Windows 
Access -- Checked

Alternatively, you could set your windows username and password in the 
Web.config file but that may not be safe. 


When the sample web service is built successfully, run QBWebConnector 
and use "Load an application" to load the .qwc files provided with this 
sample. For development purpose, you could use http://localhost... for 
your web service AppURL. We have provided a sample web connector 
configuration file (HTTPWebService.qwc) for your convenience.  You could 
also double-click on the .qwc files, which will automatically run the 
QBWebConnector loaded with the application. The .qwc (QuickBooks Web 
Connector Congfiguration) file introduces a sample web service to the 
QBWebConnector. 

This sample web service uses windows event log to trace activity while 
being updated from the QBWebConnector. To view the event log, go to 
Start -- Programs -- Administrative Tools -- Event Viewer and look under
Application section.


Building the sample
------------------
Open WCWebService.sln in Microsoft Visual Studio .NET and build the 
solution. 


Note
----
For production purpose, you would need to use https and there are some 
rules to be satisfied before QBWebConnector would load a web service. 
For information on the rules please visit QBWebConnector knowledgebase 
at: 
http://idnforums.intuit.com/messageview.aspx?catid=52&threadid=4593&enterthread=y


Useful note about using OwnerID and FileID in a real-world application
 
As part of your QB Web Connector configuration (.QWC) file, you include
OwnerID and FileID. Following note on these two parameters may be useful. 

OwnerID -- this is a GUID that represents your application or suite of 
applications, if your application needs to store private data in the 
company file for one reason or another (one of the most common cases 
being to check if you have communicated with this company file before, 
and possibly some data about that communication) that private data will 
be visible to any application that knows the OwnerID.

FileID -- this is a GUID we stamp in the file on your behalf 
(using your OwnerID) as a private data extension to the "Company" object. 
It allows an application to verify that the company file it is exchanging 
data with is consistent over time (by doing a CompanyQuery with the field 
set appropriately and reading the DataExtRet values returned.

You can find more information at the following AlphaGeek article(s): 
http://developer.intuit.com/support/technical/?id=392



