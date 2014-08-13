#pragma once

class qbXMLObject;
typedef vector<qbXMLObject*> qbXMLObjectListType;

class qbXMLObject
{
public:
    qbXMLObject() {;}

    void SetType( LPCTSTR lpszVal )        { m_strType = lpszVal; }
    void SetID( LPCTSTR lpszVal )          { m_strID = lpszVal; }
    void SetDisplayName( LPCTSTR lpszVal ) { m_strDisplayName = lpszVal; }

    string getType()        { return m_strType; }
    string getID()          { return m_strID; }
    string getDisplayName() { return m_strDisplayName; }

private:
    string m_strType;
    string m_strID;
    string m_strDisplayName;
};

class qbXMLObjectList
{
public:
    qbXMLObjectList() {;}
    ~qbXMLObjectList();

    void removeAll();

    int          getSize();
    qbXMLObject* getItemAt( long iItem );

    void addObject( qbXMLObject* pObject );
    void parseObjects( const string& strResponse, LPCTSTR lpszType );

private:
    qbXMLObject* parseObject( MSXML2::IXMLDOMNodeList* pNodeList, LPCTSTR lpszType );

    qbXMLObjectListType m_objectList;
};
