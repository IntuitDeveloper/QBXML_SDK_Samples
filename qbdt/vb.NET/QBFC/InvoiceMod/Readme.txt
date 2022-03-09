The InvoiceMod sample program allows users who are running QB 2003 R7P to modify invoices in a company file.  QuickBooks must be running and if this is the first time the sample will be run against a company file the Admin user must have the company file open.

When one runs the InvoiceMod sample, one choose at the onset to use either QBXMLRP or QBFC to be used by the sample application.  The sample continues to use that method throughout the rest of the current run of the sample program.

The first screen the sample program user is presented allows them to query for one invoice by invoice number (RefNumber) or by most of the other filter criteria (except TxnID) that can be used in an invoice query request.  MaxReturned is hard coded at 30 to prevent having to process too many invoices.  You can look at the qbXML used in the query or the QBFC calls made to create the query by checking the Show Invoice Query checkbox at the bottom left of the query screen.

Once you have performed a query you will be presented with a list of up to 30 invoices which you can choose to modify.  Select an invoice by clicking on it and then click on the Modify Invoice button to modify the invoice.

After clicking on the Modify Invoice button the InvoiceMod sample program will query for the invoice details and for list items displayed in the combo boxes on the modify invoice screen.  Depending on the size of the company file this may take a little while.  A form with invoice information will appear including a display at the bottom with invoice line information.  You may add or modify any information which can be added, removed or modified in an invoice.  When one chooses to edit an existing invoice line or add a new invoice line, an invoice line editting form will appear.  One can also delete invoice lines from an invoice.

Invoice lines appear in different colors depending on their status.  Unchanged lines from the existing invoice appear in black.  New lines added to the invoice appear in green.  Lines which have had one or more values changed appear in red.  Lines that have been deleted are shown with a line through them.

When one chooses to save one's modifications to an invoice click on the Modify Invoice button at the bottom right of the screen.  One can check the Show request check box in order to show the qbXML or the QBFC calls used to modify the invoice.