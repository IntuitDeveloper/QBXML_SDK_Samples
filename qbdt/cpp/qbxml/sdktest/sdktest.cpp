/*---------------------------------------------------------------------------
 * FILE: SDKTest.cpp
 *
 * Description:
 * QBXMLRP API tester. Run "sdktest -h" to learn how to use it.
 *
 * Created On: 09/09/2001
 *
 * Copyright (c) 2001-2013 Intuit Inc. All rights reserved.
 * Use is subject to the terms specified at:
 *     http://developer.intuit.com/legal/devsite_tos.html
 *
 */

#include <ctype.h>
#include <ctime>
#include <fstream>
#include <atlbase.h>
#include <iostream>
#include <string>

using namespace std;

#import "QBXMLRP.dll"


/* * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Prototypes
 */

void usage ();

void processRequest (CComBSTR   appName,
							CComBSTR   appID,
							TCHAR      *qbfileName,
							TCHAR      *inputXMLFileName,
							TCHAR      *outputXMLFileName);

void testVersionAndSpec (CComBSTR   appName,
                         CComBSTR   appID,
                         TCHAR      *fileName);

bool readXMLinputFile (TCHAR    *fileName,
                       BSTR     *xmlInput);

void handleError (HRESULT hr);

/* * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Definitions, Globals
 */

const int BUF_SIZE = 4096;
QBXMLRPLib::QBFileMode  qbMode = QBXMLRPLib::qbFileOpenDoNotCare;


int main (int argc, char *argv[])
{
	int         numOfParams = 1;
	char        *param = argv[1];
	TCHAR       *id = NULL;
	TCHAR       *name = NULL;
	CComBSTR    appID;
	CComBSTR    appName;
	bool        showUsage = false;
	bool        testVersion = false;

  if (argc == 1  ||  param == 0) {
    std::cout << "\nFor help, use \"-h\" option.\n";
    return 0;
  } else {
    int     i;
    char    ch;

    for (i = 1; i < argc; i++) {
      param = argv [i];
      if (*param == '-') {
        ch = tolower (*++param);

        switch (ch) {
          case 'h':
            showUsage = true;
            break;
          case 'v':
            testVersion = true;
            numOfParams += 1;
            break;
          case 'n':
            if (++i < argc) {
              name = argv [i];
              numOfParams += 2;
            } else {
              showUsage = true;
            }
            break;
          case 'i':
            if (++i < argc) {
              id = argv [i];
              numOfParams += 2;
            } else {
              showUsage = true;
            }
            break;
          case 'o':
            if (++i < argc) {
              char  *mode = argv [i];
              if (strlen (mode) == 2  &&  tolower (*(mode + 1)) == 'u') {
                if (tolower (*mode) == 'm') {
                  qbMode = QBXMLRPLib::qbFileOpenMultiUser;
                } else if (tolower (*mode) == 's') {
                  qbMode = QBXMLRPLib::qbFileOpenSingleUser;
                } else {
                  showUsage = true;
                }
                numOfParams += 2;
              } else {
                showUsage = true;
              }
            } else {
              showUsage = true;
            }
            break;
          default:
            std::cout << "\nFor help, use \"-h\" option\n";
            return 0;
        }
      }
    }
  }

  if (showUsage  ||  (!testVersion  &&  argc != numOfParams + 3)) {
	  usage ();
	  return 0;
  }

  if (name == NULL) {
    name = "SdkTest";
  }
  appName = name;
  appID	 = id;

  CoInitialize (NULL);

  if (testVersion) {
    testVersionAndSpec (appName, appID, argv [argc - 1]);
  } else {
    processRequest (appName, appID, argv [argc - 3], argv [argc - 2], argv [argc - 1]);
  }

  CoUninitialize ();

  return 0;

}

void usage () {
	std::cout << "SDKTest v. 1.07\n";
	std::cout << "Usage:\n";
	std::cout << "\tsdktest [-h]  or\n";
	std::cout << "\tsdktest [-o MU/SU] [-n appName] [-i appID] qb_file input_file output_file\n";
	std::cout << "\tsdktest [-v] qb_file\n";
	std::cout << "\n";
	std::cout << "where:\n";
	std::cout << "\t-h to see this help\n";
	std::cout << "\t-o to open the company data file in the MultiUser or SingleUser mode\n";
	std::cout << "\t   (\"DoNotCare\" mode is used if this option is not specified)\n";
	std::cout << "\t-n appName provides the application name\n";
	std::cout << "\t-i appID provides the application ID\n\n";
	std::cout << "\t-v to test the version interfaces\n";
	std::cout << "\tqb file is a QuickBooks company data file name or \"\" to send the request to the currently open file\n";
	std::cout << "\tinput file has to exist and contains the XML data to be tested\n";
	std::cout << "\toutput file will be created and will contain an XML response for the given request.\n";
}


/*---------------------------------------------------------------------------
 * processRequest
 *
 * 1. read the XML data file
 * 2. instantiate the RequestProcessor smart pointer
 * 3. Open Connection, app name and id provided as parameters or default
 * 4. Begin Session, the QuickBooks company data file is provided as a
 * parameter, the mode is provided or as a parameter or "DoNotCare" is a default
 * mode, remember the ticket
 * 5. show the currently open data file
 * 6. Process the Request, use the ticket and data read from the provided XML
 * data file,
 * 7. dump the response
 * 8. End the Session and Close the Connection
 *
 */

void processRequest (CComBSTR   appName,
							CComBSTR   appID,
							TCHAR      *fileName,
							TCHAR      *inputXMLFileName,
							TCHAR      *outputXMLFileName)
{
  CComBSTR  qbFile (fileName);
  BSTR      xmlInput = NULL;
  BSTR      xmlOutput = NULL;
  BSTR      companyName = NULL;
   
  clock_t     startTime;
  clock_t     finishTime;
  double      duration;
  TCHAR       durationStr [100];
  HRESULT     hr = S_OK;
  BSTR        ticket = NULL;
  QBXMLRPLib::IRequestProcessor2Ptr pReqProc;

  if (!readXMLinputFile (inputXMLFileName, &xmlInput)) {
    std::cout << "SDKTEST: Could not read the XML input file.\n";
	 return;
  }

  try {
    std::cout << "SDKTEST: COM access to SDK starting ...\n";
    // start the clock
    startTime = clock();
    
    hr = pReqProc.CreateInstance (__uuidof (QBXMLRPLib::RequestProcessor));
    if (SUCCEEDED (hr)  &&  pReqProc != NULL) {
      hr = pReqProc->raw_OpenConnection (appID.m_str, appName.m_str);
      if (SUCCEEDED (hr)) {

        hr = pReqProc->raw_BeginSession (qbFile.m_str, qbMode, &ticket);
        if (SUCCEEDED (hr)) {
          hr = pReqProc->raw_GetCurrentCompanyFileName (ticket, &companyName);
          if (SUCCEEDED (hr)) {
            std::cout << "data file name: " << _com_util::ConvertBSTRToString (companyName) << "\n";
          } else {
            std::cout << "could not get the data file name\n";
          }

          hr = pReqProc->raw_ProcessRequest (ticket, xmlInput, &xmlOutput);
          if (SUCCEEDED (hr)) {
            std::cout << "process request OK, dumping the response.\n";
            // dump the response
            std::ofstream   outputFile (outputXMLFileName);
            if (outputFile == 0) {
              std::cout << "could not open output file" << outputXMLFileName << "\n";
            } else {
              TCHAR *outputData2;   // this pointer does not change;
              TCHAR *outputData = _com_util::ConvertBSTRToString (xmlOutput);
              outputData2 = outputData;   // keep it so that we can delete it
              if (outputData != NULL) {
                while (*outputData != 0) {
                  outputFile.put (*outputData++);
                }
              }
              outputFile.close ();
              delete [] outputData2;
              SysFreeString (xmlOutput);
            }
          } else {
            std::cout << "process request failed\n";
          }
          pReqProc->raw_EndSession (ticket);
        }
        pReqProc->raw_CloseConnection ();
      }
    }
    finishTime = clock();

    if (FAILED (hr)) {
      handleError (hr);
    }
    duration = (double) (finishTime - startTime) / CLOCKS_PER_SEC;
    sprintf (durationStr, "SDKTest: Finished after %3.3f seconds\n", duration);
    std::cout << durationStr;
  } catch (const _com_error& err) {
    std::cout << "Error: ";
    _bstr_t desc = err.Description();
    if (desc.length() > 0) {
      std::cout << (char *) desc;
    } else {
      const TCHAR * message = err.ErrorMessage();
      if (message != NULL) {
        std::cout << message;
      } else {
        std::cout << "No additional information";
      }
    }
  }
 
}   // processRequest


/*---------------------------------------------------------------------------
 * testVersionAndSpec
 *
 * tests all the version and release methods and QBXMLVersionsForSession method.
 *
 */

void testVersionAndSpec (CComBSTR   appName,
                         CComBSTR   appID,
                         TCHAR      *fileName)
{
  HRESULT     hr = S_OK;
  CComBSTR    qbFile (fileName);
  BSTR        companyName = NULL;
  BSTR        ticket = NULL;

  QBXMLRPLib::IRequestProcessor2Ptr pReqProc;

  hr = pReqProc.CreateInstance (__uuidof (QBXMLRPLib::RequestProcessor));

  if (SUCCEEDED (hr)  &&  pReqProc != NULL) {
    hr = pReqProc->raw_OpenConnection (appID.m_str, appName.m_str);
    if (SUCCEEDED (hr)) {
      hr = pReqProc->raw_BeginSession (qbFile.m_str, qbMode, &ticket);
      if (SUCCEEDED (hr)) {
        hr = pReqProc->raw_GetCurrentCompanyFileName (ticket, &companyName);
        if (SUCCEEDED (hr)) {
          std::cout << "data file name: " << _com_util::ConvertBSTRToString (companyName) << "\n\n";
        } else {
          std::cout << "could not get the data file name\n";
        }

        short         majorVersion;
        short         minorVersion;
        short         releaseNumber;
        char          releaseLevelChar;
        TCHAR         versionBuffer [20];
        SAFEARRAY     *psa;
        QBXMLRPLib::QBXMLRPReleaseLevel releaseLevel;

        hr = pReqProc->get_MajorVersion (&majorVersion);
        if (SUCCEEDED (hr)) {
          hr = pReqProc->get_MinorVersion (&minorVersion);
        }
        if (SUCCEEDED (hr)) {
          hr = pReqProc->get_ReleaseLevel (&releaseLevel);
        }
        if (SUCCEEDED (hr)) {
          hr = pReqProc->get_ReleaseNumber (&releaseNumber);
        }
        if (SUCCEEDED (hr)) {
          hr = pReqProc->get_QBXMLVersionsForSession (ticket, &psa);
        }
        if (SUCCEEDED (hr)) {
          BSTR * bstrArray;

          switch (releaseLevel) {
            case QBXMLRPLib::alpha:
              releaseLevelChar = 'A';
              break;
            case QBXMLRPLib::beta:
              releaseLevelChar = 'B';
              break;
            case QBXMLRPLib::release:
              releaseLevelChar = 'R';
              break;
            case QBXMLRPLib::preAlpha:
            default:
              releaseLevelChar = 'P';
              break;
          }
          sprintf (versionBuffer, "%d.%d %c%d",
                   majorVersion, minorVersion,
                   releaseLevelChar, releaseNumber);

          std::cout << "QBXMLRP Request Processor version: " << versionBuffer << "\n"; 

          SafeArrayAccessData(psa, (void **) &bstrArray);
          for (UINT i = 0; i < psa->rgsabound->cElements; i++) {
            std::cout << "qbXML spec versions supported:     " << _com_util::ConvertBSTRToString (bstrArray[i]) << "\n"; 
          }
          SafeArrayUnaccessData(psa);
        } else {
          std::cout << "No supported versions available! Check for missing DTDs\n";
        }
        pReqProc->raw_EndSession (ticket);
      }
      pReqProc->raw_CloseConnection ();
    }
  }

  if (FAILED (hr)) {
    handleError (hr);
  }

}   // testVersionAndSpec


/*---------------------------------------------------------------------------
 * readXMLInputFile
 *
 * opens the input XML data file, reads the data, creates a BSTR
 *
 */

bool readXMLinputFile (TCHAR    *fileName,
                       BSTR     *pXmlInput)
{
  std::ifstream   inputFile (fileName);

  if (inputFile == NULL) {
    return false;
  }

  TCHAR       buffer [BUF_SIZE + 1];
  TCHAR       ch;
  int         count;
  string      stringRequest;

  while (true) {
    count = 0;

    while (count < BUF_SIZE  &&  inputFile.get (ch)) {
      buffer [count++] = ch;
    }

    // any data?
    if (count == 0) {
      return false;
    }

    buffer [count] = 0;

    stringRequest.append (buffer);

    if (count < BUF_SIZE) {
      break;
    }
  }
  inputFile.close ();

  CComBSTR  xmlInput = CComBSTR (stringRequest.c_str());
  *pXmlInput = xmlInput.Detach ();
  return true;

}   // readXMLinputFile


/*---------------------------------------------------------------------------
 * handleError
 *
 * using the ErrorInfo interface, get the detailed information about the provided HRESULT.
 *
 */

void handleError (HRESULT hr)
{
  if (FAILED (hr)) {
    TCHAR         errorStr [2048];
    IErrorInfo    *errorPtr;
    HRESULT       errorHr;

    errorHr = GetErrorInfo (0, &errorPtr);
    if (errorHr == S_OK  &&  errorPtr != NULL) {
      BSTR    description;
      TCHAR   *pDescription;
      try {
        errorPtr->GetDescription (&description);
        pDescription = _com_util::ConvertBSTRToString(description);
        sprintf (errorStr, "\terror = %lx, %s\n", hr, pDescription);
        delete [] pDescription;
      } catch (const _com_error& ex) {
        ex.Error();
        sprintf (errorStr, "\terror = %lx, %s\n", hr, "error getting error text.");
      }
    } else {
      sprintf (errorStr, "\terror = %lx\n", hr);
    }
    std::cout << errorStr;
  }

}   // handleError

/*
 * End of SDKTest.cpp
 */
