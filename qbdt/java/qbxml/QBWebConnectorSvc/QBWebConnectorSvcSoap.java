/**
 * QBWebConnectorSvcSoap.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis 1.2.1 Jun 14, 2005 (09:15:57 EDT) WSDL2Java emitter.
 */

package com.intuit.developer;

public interface QBWebConnectorSvcSoap extends java.rmi.Remote {
    public com.intuit.developer.AuthResponse authenticate2(java.lang.String strUserName, java.lang.String strPassword) throws java.rmi.RemoteException;
    public com.intuit.developer.ArrayOfString authenticate(java.lang.String strUserName, java.lang.String strPassword) throws java.rmi.RemoteException;
    public java.lang.String sendRequestXML(java.lang.String ticket, java.lang.String strHCPResponse, java.lang.String strCompanyFileName, java.lang.String qbXMLCountry, int qbXMLMajorVers, int qbXMLMinorVers) throws java.rmi.RemoteException;
    public int receiveResponseXML(java.lang.String ticket, java.lang.String response, java.lang.String hresult, java.lang.String message) throws java.rmi.RemoteException;
    public java.lang.String connectionError(java.lang.String ticket, java.lang.String hresult, java.lang.String message) throws java.rmi.RemoteException;
    public java.lang.String getLastError(java.lang.String ticket) throws java.rmi.RemoteException;
    public java.lang.String closeConnection(java.lang.String ticket) throws java.rmi.RemoteException;
}
