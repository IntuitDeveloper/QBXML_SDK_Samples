/**
 * AuthResponse.java
 *
 * This file was auto-generated from WSDL
 * by the Apache Axis 1.2.1 Jun 14, 2005 (09:15:57 EDT) WSDL2Java emitter.
 */

package com.intuit.developer;

public class AuthResponse  implements java.io.Serializable {
    private java.lang.String sessionTicket;
    private java.lang.String companyFile;
    private java.lang.String authResult;

    public AuthResponse() {
    }

    public AuthResponse(
           java.lang.String sessionTicket,
           java.lang.String companyFile,
           java.lang.String authResult) {
           this.sessionTicket = sessionTicket;
           this.companyFile = companyFile;
           this.authResult = authResult;
    }


    /**
     * Gets the sessionTicket value for this AuthResponse.
     * 
     * @return sessionTicket
     */
    public java.lang.String getSessionTicket() {
        return sessionTicket;
    }


    /**
     * Sets the sessionTicket value for this AuthResponse.
     * 
     * @param sessionTicket
     */
    public void setSessionTicket(java.lang.String sessionTicket) {
        this.sessionTicket = sessionTicket;
    }


    /**
     * Gets the companyFile value for this AuthResponse.
     * 
     * @return companyFile
     */
    public java.lang.String getCompanyFile() {
        return companyFile;
    }


    /**
     * Sets the companyFile value for this AuthResponse.
     * 
     * @param companyFile
     */
    public void setCompanyFile(java.lang.String companyFile) {
        this.companyFile = companyFile;
    }


    /**
     * Gets the authResult value for this AuthResponse.
     * 
     * @return authResult
     */
    public java.lang.String getAuthResult() {
        return authResult;
    }


    /**
     * Sets the authResult value for this AuthResponse.
     * 
     * @param authResult
     */
    public void setAuthResult(java.lang.String authResult) {
        this.authResult = authResult;
    }

    private java.lang.Object __equalsCalc = null;
    public synchronized boolean equals(java.lang.Object obj) {
        if (!(obj instanceof AuthResponse)) return false;
        AuthResponse other = (AuthResponse) obj;
        if (obj == null) return false;
        if (this == obj) return true;
        if (__equalsCalc != null) {
            return (__equalsCalc == obj);
        }
        __equalsCalc = obj;
        boolean _equals;
        _equals = true && 
            ((this.sessionTicket==null && other.getSessionTicket()==null) || 
             (this.sessionTicket!=null &&
              this.sessionTicket.equals(other.getSessionTicket()))) &&
            ((this.companyFile==null && other.getCompanyFile()==null) || 
             (this.companyFile!=null &&
              this.companyFile.equals(other.getCompanyFile()))) &&
            ((this.authResult==null && other.getAuthResult()==null) || 
             (this.authResult!=null &&
              this.authResult.equals(other.getAuthResult())));
        __equalsCalc = null;
        return _equals;
    }

    private boolean __hashCodeCalc = false;
    public synchronized int hashCode() {
        if (__hashCodeCalc) {
            return 0;
        }
        __hashCodeCalc = true;
        int _hashCode = 1;
        if (getSessionTicket() != null) {
            _hashCode += getSessionTicket().hashCode();
        }
        if (getCompanyFile() != null) {
            _hashCode += getCompanyFile().hashCode();
        }
        if (getAuthResult() != null) {
            _hashCode += getAuthResult().hashCode();
        }
        __hashCodeCalc = false;
        return _hashCode;
    }

    // Type metadata
    private static org.apache.axis.description.TypeDesc typeDesc =
        new org.apache.axis.description.TypeDesc(AuthResponse.class, true);

    static {
        typeDesc.setXmlType(new javax.xml.namespace.QName("http://developer.intuit.com/", "AuthResponse"));
        org.apache.axis.description.ElementDesc elemField = new org.apache.axis.description.ElementDesc();
        elemField.setFieldName("sessionTicket");
        elemField.setXmlName(new javax.xml.namespace.QName("http://developer.intuit.com/", "SessionTicket"));
        elemField.setXmlType(new javax.xml.namespace.QName("http://www.w3.org/2001/XMLSchema", "string"));
        elemField.setMinOccurs(0);
        elemField.setNillable(false);
        typeDesc.addFieldDesc(elemField);
        elemField = new org.apache.axis.description.ElementDesc();
        elemField.setFieldName("companyFile");
        elemField.setXmlName(new javax.xml.namespace.QName("http://developer.intuit.com/", "CompanyFile"));
        elemField.setXmlType(new javax.xml.namespace.QName("http://www.w3.org/2001/XMLSchema", "string"));
        elemField.setMinOccurs(0);
        elemField.setNillable(false);
        typeDesc.addFieldDesc(elemField);
        elemField = new org.apache.axis.description.ElementDesc();
        elemField.setFieldName("authResult");
        elemField.setXmlName(new javax.xml.namespace.QName("http://developer.intuit.com/", "AuthResult"));
        elemField.setXmlType(new javax.xml.namespace.QName("http://www.w3.org/2001/XMLSchema", "string"));
        elemField.setMinOccurs(0);
        elemField.setNillable(false);
        typeDesc.addFieldDesc(elemField);
    }

    /**
     * Return type metadata object
     */
    public static org.apache.axis.description.TypeDesc getTypeDesc() {
        return typeDesc;
    }

    /**
     * Get Custom Serializer
     */
    public static org.apache.axis.encoding.Serializer getSerializer(
           java.lang.String mechType, 
           java.lang.Class _javaType,  
           javax.xml.namespace.QName _xmlType) {
        return 
          new  org.apache.axis.encoding.ser.BeanSerializer(
            _javaType, _xmlType, typeDesc);
    }

    /**
     * Get Custom Deserializer
     */
    public static org.apache.axis.encoding.Deserializer getDeserializer(
           java.lang.String mechType, 
           java.lang.Class _javaType,  
           javax.xml.namespace.QName _xmlType) {
        return 
          new  org.apache.axis.encoding.ser.BeanDeserializer(
            _javaType, _xmlType, typeDesc);
    }

}
