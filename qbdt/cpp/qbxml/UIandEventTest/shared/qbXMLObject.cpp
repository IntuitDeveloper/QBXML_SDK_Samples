
#include "stdafx.h"
#include "qbXMLObject.h"
#include "Utility.h"

/////////////////////////////////////////////////////////////////////
// qbXMLObject


/////////////////////////////////////////////////////////////////////
// qbXMLObjectList

//qbXMLObjectList destructor.
qbXMLObjectList::~qbXMLObjectList() {
    removeAll();
}

// removeAll
// Removes all items from list..
void qbXMLObjectList::removeAll() {
  qbXMLObjectListType::iterator it;
  for (it = m_objectList.begin(); it != m_objectList.end(); it++) {
    qbXMLObject *pObject = *it;
    delete pObject;
  }
  m_objectList.clear();
}

// getSize
// Returns number of items in list.
int qbXMLObjectList::getSize() {
    return m_objectList.size();
}

// getItemAt
// Return pointer to item at given position.
qbXMLObject* qbXMLObjectList::getItemAt( long iItem ) {
    qbXMLObject* pObject = NULL;
    if( iItem >= 0 && iItem < (int)m_objectList.size() )  {
        pObject = m_objectList.at( iItem );
    }
    return pObject;
}

// addObject
// Adds an object to object list.
void qbXMLObjectList::addObject( qbXMLObject* pObject ) {
    m_objectList.push_back( pObject );
}


// parseObjects
// Load objects from response XML.
void qbXMLObjectList::parseObjects( const string& strResponse,
                                    LPCTSTR lpszType) {
    // Parse response into a DOM object.
    string strReason;
    MSXML2::IXMLDOMDocument *pXMLDocPtr = Utility::LoadXML( strResponse, strReason );
    if( pXMLDocPtr == NULL ) {
        string strMessage = "Error parsing response XML.\n\n";
        strMessage += strReason;
        AfxMessageBox( strMessage.c_str() );
        return;
    }

    // Get data "ret" objects.
    MSXML2::IXMLDOMNodeList* pResultList = NULL;
    CComBSTR tagName( lpszType );
    tagName.Append( "Ret" );
    HRESULT hr = pXMLDocPtr->getElementsByTagName( tagName, &pResultList );
    if( FAILED(hr) || pResultList == NULL ) {
        AfxMessageBox( "Error parsing response XML.\n\nUnable to find response aggregate." );
        return;
    }
    
    // Get count of data "ret" objects.
    long tagCount = 0;
    hr = pResultList->get_length( &tagCount );

    // Process data "ret" objects.
    for( int i = 0; i < tagCount; i++ ) {
        // Get this response object.
        MSXML2::IXMLDOMNode *pListItem = NULL;
        hr = pResultList->get_item( i, &pListItem );
        if( SUCCEEDED(hr) && pListItem ) {
            // Get node list of this response object.
            MSXML2::IXMLDOMNodeList *pNodeList = NULL;
            hr = pListItem->get_childNodes( &pNodeList );
            if( SUCCEEDED(hr) && pNodeList ) {
                // Build an object from node list.
                qbXMLObject* pObject = parseObject( pNodeList, lpszType );
                // If we got an object, add it to object list.
                if( pObject ) {
                    addObject( pObject );
                }
            }
        }
    }

    return;
}

// parseObject
// Load an object from response XML.
// Returns object.
qbXMLObject* qbXMLObjectList::parseObject( MSXML2::IXMLDOMNodeList* pNodeList,
                                           LPCTSTR lpszType ) {
    HRESULT hr = S_OK;
    
    // Get number of nodes.
    long nodeCount;
    hr = pNodeList->get_length( &nodeCount );
    if( FAILED(hr) || nodeCount == 0 ) {
        return NULL;
    }
    
    string strID, strDisplayName;

    // Go through nodes.
    MSXML2::IXMLDOMNode *pNode = NULL;
    for( int i = 0; i < nodeCount; i++ ) {
        // Get this node.
        hr = pNodeList->get_item( i, &pNode );
        if( SUCCEEDED(hr) && pNode ) {
            // Get name and value of node.   
            BSTR bstrName, bstrValue;
            hr = pNode->get_nodeName( &bstrName );
            if( SUCCEEDED(hr) ) {
                hr = pNode->get_text( &bstrValue );
            }
            // If we got node info . . .
            if( SUCCEEDED(hr) ) {
                string strName, strValue;
                Utility::BSTRToString( bstrName, strName );
                Utility::BSTRToString( bstrValue, strValue );

                // If node is one we're interested in, save its value.
                if( strName.compare("TxnID") == 0 ) {
                    strID = strValue;
                }
                else if( strName.compare("ListID") == 0 ) {
                    strID = strValue;
                }
                else if( strName.compare("RefNumber") == 0 ) {
                    strDisplayName = strValue;
                }
                else if( strName.compare("FullName") == 0 ) {
                    strDisplayName = strValue;
                }
                else if( strName.compare("Name") == 0 ) {
                    if( strDisplayName.length() == 0 ) {
                        strDisplayName = strValue;
                    }
                }
            }
        }
    }

    // If we got an ID and name for object, create the object.
    qbXMLObject* pObject = NULL;
    if( strID.length() > 0 ) {
        pObject = new qbXMLObject;
        pObject->SetType( lpszType );
        pObject->SetID( strID.c_str() );
        if( strDisplayName.length() > 0 ) {
            pObject->SetDisplayName( strDisplayName.c_str() );
        }
        else {
            pObject->SetDisplayName( strID.c_str() );
        }
    }

    return pObject;
}
