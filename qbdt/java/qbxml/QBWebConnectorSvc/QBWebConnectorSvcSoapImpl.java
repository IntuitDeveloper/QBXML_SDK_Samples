/**
 * QBWebConnectorSvcSoapImpl.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis 1.2.1 Jun 14, 2005 (09:15:57 EDT) WSDL2Java emitter.
 */

package com.intuit.developer;

public class QBWebConnectorSvcSoapImpl implements com.intuit.developer.QBWebConnectorSvcSoap{
    public com.intuit.developer.AuthResponse authenticate2(java.lang.String strUserName, java.lang.String strPassword) throws java.rmi.RemoteException {
        return null;
    }

    public com.intuit.developer.ArrayOfString authenticate(java.lang.String strUserName, java.lang.String strPassword) throws java.rmi.RemoteException
    {
	    String[] asRtn = new String[2];
	    asRtn[0] = "{F5FCCBC3-AA13-4d28-9DBF-3E571823F2BB}"; //myGUID.toString();
	    asRtn[1] = "none";
	    System.out.println("In authenticate new two");
	    ArrayOfString asRtn2 = new ArrayOfString(asRtn);
	    System.out.println("In authenticate step2");
	    System.out.println("In authenticate as[0] = " + asRtn2.getString(0));
	    System.out.println("In authenticate as[1] = " + asRtn2.getString(1));
	    return asRtn2;
    }

    public java.lang.String sendRequestXML(java.lang.String ticket, java.lang.String strHCPResponse, java.lang.String strCompanyFileName, java.lang.String qbXMLCountry, int qbXMLMajorVers, int qbXMLMinorVers) throws java.rmi.RemoteException {
        return null;
    }

    public int receiveResponseXML(java.lang.String ticket, java.lang.String response, java.lang.String hresult, java.lang.String message) throws java.rmi.RemoteException {
        return -3;
    }

    public java.lang.String connectionError(java.lang.String ticket, java.lang.String hresult, java.lang.String message) throws java.rmi.RemoteException {
        return null;
    }

    public java.lang.String getLastError(java.lang.String ticket) throws java.rmi.RemoteException {
        return null;
    }

    public java.lang.String closeConnection(java.lang.String ticket) throws java.rmi.RemoteException {
        System.out.println("In closeConnection");
        System.out.println(ticket);
        return("close with this message");
    }

}
