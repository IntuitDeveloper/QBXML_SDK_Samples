There are three components to this sample. 

- Client side demonstrates what an end user would see when they access the application. 
- Merchant side demonstrates how a merchant can setup their web-connector to point to the web-service.
- Service side demonstrates how a service provides data correspondence with QuickBooks.

These three components are listed below in order of execution of steps: -

PREPARATION
1) Copy the entire WCECommSample directory into your c:\inetup\wwwroot path. If you are using Microsoft Developer Studio then the tool should automatically do this step for you. 

2) From IIS_Console, right-click WCECommSample (client, merchant, service) -> Properties -> click Create button next to "Application Name". 

CLIENT SIDE TRANSACTION
3) Point your IE browser to http://localhost/WCECommSample/client/Default.aspx 

4) Database ecommdb is pre-populated with some username/password (encryption not used for simplicity of the sample) used 
to log on to the client site. If you want to create a new user, use "sign-up" from log on page. New user created here, will be transfered as Customer into QuickBooks on next update session with web connector.

5) Database ecommdb is also pre-populated with some sample items (phone calling cards) that you could use for this sample. On the next update session with web connector (via Merchant site) any inventory items that you have on QB will be transfered to ecommdb. 

6) After you log on, do some purchasing (it's really a simulation and not tied to any payment system, so don't worry :-)) on client to create some txns (salesreceipts)


MERCHANT SITE ADMINISTRATION
7) Point your IE browser to http://localhost/WCECommSample/merchant/setup.html

8) Start with an empty company file (no customers and inventory items).

9) Setup WCECommService with QBWC using merchant site and following the steps described in the site.


WEB SERVICE 
When you run update from QBWC (directly or via merchant site administration page), QBWC exchanges data with this web service. This web service provides code to insert, update or retrieve data from ecommdb. 


