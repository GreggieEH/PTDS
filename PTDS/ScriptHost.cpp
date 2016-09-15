#include "stdafx.h"
#include "ScriptHost.h"
#include "MyObject.h"
#include "ScriptSite.h"

CScriptHost::CScriptHost(CMyObject * pObject) :
	m_pObject(pObject),
	m_pActiveScript(NULL),
	m_pdispHost(NULL)			// host object
{
	this->m_szHost[0] = '\0';
}

CScriptHost::~CScriptHost()
{
	this->UnloadScript();
}

BOOL CScriptHost::GetScriptLoaded()
{
	SCRIPTSTATE			ss;			// script state
	BOOL				fScriptLoaded = FALSE;

	if (NULL != this->m_pActiveScript)
	{
		this->m_pActiveScript->GetScriptState(&ss);
		fScriptLoaded = SCRIPTSTATE_STARTED == ss || SCRIPTSTATE_CONNECTED == ss;
	}
	return fScriptLoaded;
}

// get the script engine
HRESULT CScriptHost::GetScriptEngine(
						LPCTSTR		szLanguage,
						CLSID	*	pClsid)
{
	HRESULT				hr;
	ICatInformation	*	pCatInformation;
	CATID				catid;
	IEnumCLSID		*	pEnumCLSID = NULL;			// enumeration of the scripting engines
	BOOL				fDone;				// completed flag
	CLSID				clsid;				// member of the enumeration
	LPOLESTR			UserType	= NULL;			// user type for the scripting engine
//	TCHAR				szUserType[MAX_PATH];

	*pClsid = CLSID_NULL;
	// use the ICatInformation object to enumerate the ActiveScript parsing engines
	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatInformation, (LPVOID*)&pCatInformation);
	if (SUCCEEDED(hr))
	{
		catid = CATID_ActiveScriptParse;
		hr = pCatInformation->EnumClassesOfCategories(1, &catid, -1, NULL, &pEnumCLSID);
		pCatInformation->Release();
	}
	// find the class id for the JScript parser
	if (SUCCEEDED(hr))
	{
		fDone = FALSE;
		while (!fDone && S_OK == pEnumCLSID->Next(1, &clsid, NULL))
		{
			// check if this is a JScript clsid
			UserType = NULL;
			hr = OleRegGetUserType(clsid, USERCLASSTYPE_FULL, &UserType);
			if (SUCCEEDED(hr))
			{
				//WideCharToMultiByte(CP_ACP, 0, UserType, -1, szUserType, MAX_PATH, NULL, NULL);
				//CoTaskMemFree((LPVOID)UserType);
				fDone = NULL != StrStrI(UserType, szLanguage);
				if (fDone)
				{
					// store the class id
					*pClsid = clsid;
				}
	//			CoTaskMemFree((LPVOID)szUserType);
				CoTaskMemFree((LPVOID)UserType);
			}
		}
		pEnumCLSID->Release();
	}
	return fDone ? S_OK : E_FAIL;
}

HRESULT CScriptHost::LoadScript(
				LPCTSTR			szFilePath,
				IDispatch	*	pdispHost,
				LPCTSTR			szHost)
{
	HRESULT					hr;
	CScriptSite			*	pScriptSite = NULL;
	IActiveScriptSite	*	pActiveScriptSite = NULL;
	IActiveScriptParse	*	pActiveScriptParse = NULL;
	CLSID					clsidScriptEngine;
	EXCEPINFO				ei;								// exception information
//	IUnknown			*	punkScriptSite = NULL;			// script site
	LPOLESTR				pOleStrScript	= NULL;			// script string
//	WCHAR					pHost[MAX_PATH];
	IActiveScriptProperty*	pScriptProperty = NULL;			// active script property

	// obtain the scripting engine
	hr = this->GetScriptEngine(TEXT("JScript"), &clsidScriptEngine);
	if (SUCCEEDED(hr))
	{
		hr = CoCreateInstance(clsidScriptEngine, 0, CLSCTX_ALL, IID_IActiveScriptParse, (void**)&pActiveScriptParse);
	}
	if (SUCCEEDED(hr))
	{
		hr = pActiveScriptParse->QueryInterface(IID_IActiveScript, (void**)&(this->m_pActiveScript));
	}
	// set script state to INITIALIZED 
	if (SUCCEEDED(hr))
		hr = pActiveScriptParse->InitNew();
	if (SUCCEEDED(hr))
	{
		// set the script property
		hr = this->m_pActiveScript->QueryInterface(IID_IActiveScriptProperty, (LPVOID*)&pScriptProperty);
		if (SUCCEEDED(hr))
		{
			VARIANT		Value;
			InitVariantFromInt32(SCRIPTLANGUAGEVERSION_5_8, &Value);
			hr = pScriptProperty->SetProperty(SCRIPTPROP_INVOKEVERSIONING, NULL, &Value);
			pScriptProperty->Release();
		}

		// create the script site
		pScriptSite = new CScriptSite(this);
		if (pScriptSite)
		{
			pScriptSite->Init();
			hr = pScriptSite->QueryInterface(IID_IActiveScriptSite, (LPVOID*)&pActiveScriptSite);
		}
		else
			hr = E_OUTOFMEMORY;
		if (SUCCEEDED(hr))
		{
			hr = this->m_pActiveScript->SetScriptSite(pActiveScriptSite);
			pActiveScriptSite->Release();
		}
	}
	// add host shell objects to engine's namespace and set state to STARTED
	if (SUCCEEDED(hr) && NULL != szHost)
	{
		StringCchCopy(this->m_szHost, MAX_PATH, szHost);
		this->m_pdispHost = pdispHost;
		this->m_pdispHost->AddRef();
	//	SHAnsiToUnicode(this->m_szHost, pHost, MAX_PATH);
		hr = this->m_pActiveScript->AddNamedItem(this->m_szHost, SCRIPTITEM_ISVISIBLE);
		if (SUCCEEDED(hr)) hr = this->m_pActiveScript->SetScriptState(SCRIPTSTATE_STARTED);
	}
	// add supplied script files to the engine's parsed state
	if (SUCCEEDED(hr))
	{
		ZeroMemory(&ei, sizeof(ei));
		if (this->LoadScriptFile(szFilePath, &pOleStrScript))
		{
			hr = pActiveScriptParse->ParseScriptText(
				pOleStrScript,
				0, 0, 0, 0, 0, SCRIPTTEXT_ISVISIBLE, 0, &ei);
			CoTaskMemFree((LPVOID)pOleStrScript);
		}
		else hr = E_FAIL;
	}
	Utils_RELEASE_INTERFACE(pActiveScriptParse);
	// force the engine to connect any outbound interfaces to the 
	// host's objects
	if (SUCCEEDED(hr))
		hr = this->m_pActiveScript->SetScriptState(SCRIPTSTATE_CONNECTED);
	else
	{
		this->UnloadScript();
	}
	return hr;
}

void CScriptHost::UnloadScript()
{
	Utils_RELEASE_INTERFACE(this->m_pdispHost);
	this->m_szHost[0] = '\0';

	if (NULL != this->m_pActiveScript)
	{
		m_pActiveScript->SetScriptState(SCRIPTSTATE_DISCONNECTED);
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		Utils_RELEASE_INTERFACE(this->m_pActiveScript);
	}
}

// get main window
HWND CScriptHost::GetMainWindow()
{
	HWND		hwndMain = NULL;
//	this->m_pObject->FireRequestMainWindow(&hwndMain);
	return hwndMain;
}
// script error
void CScriptHost::OnScriptError(
	LPCTSTR			szError)
{
	MessageBox(this->GetMainWindow(), szError, TEXT("SciInput"), MB_OK);
}

// get script dispatch
BOOL CScriptHost::GetScriptDispatch(
	IDispatch	**	ppdispScript)
{
	HRESULT			hr;

	if (NULL != this->m_pActiveScript)
	{
		hr = this->m_pActiveScript->GetScriptDispatch(NULL, ppdispScript);
		return SUCCEEDED(hr);
	}
	else
	{
		*ppdispScript = NULL;
		return FALSE;
	}
}

BOOL CScriptHost::LoadScriptFile(
	LPCTSTR		szFilePath,
	LPOLESTR*	pszScript)
{
	//	HRESULT			hr;
	HANDLE			hFile;
	BOOL			fSuccess = FALSE;
	DWORD			fileSize;
	BYTE		*	szFileBuffer = NULL;
	DWORD			dwNRead;
	DWORD			fileOffset = 0;

	*pszScript = NULL;

	hFile = CreateFile(szFilePath, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		fileSize = GetFileSize(hFile, NULL);
		szFileBuffer = (LPBYTE)CoTaskMemAlloc(fileSize + sizeof(TCHAR));
		ZeroMemory((PVOID)szFileBuffer, fileSize + sizeof(TCHAR));
		if (ReadFile(hFile, (LPVOID)szFileBuffer, fileSize, &dwNRead, NULL))
		{
			// endianness flag
			if (255 == szFileBuffer[0] && 254 == szFileBuffer[1])
			{
				fileOffset = 2;
			}
			if (IsTextUnicode((const void*)&(szFileBuffer[fileOffset]), fileSize, 0))
			{
				*pszScript = (LPOLESTR)CoTaskMemAlloc(fileSize);
				ZeroMemory((PVOID)*pszScript, fileSize);
				CopyMemory((PVOID)*pszScript, &(szFileBuffer[fileOffset]), fileSize - fileOffset);
				fSuccess = TRUE;
			}
			else
			{
				// convert to unicode
				int			nChars = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)szFileBuffer, -1, NULL, 0);
				*pszScript = (LPOLESTR)CoTaskMemAlloc((nChars + 1) * sizeof(OLECHAR));
				MultiByteToWideChar(CP_ACP, 0, (LPCSTR)szFileBuffer, -1, *pszScript, nChars + 1);
				fSuccess = TRUE;
			}
		}
		CoTaskMemFree((LPVOID)szFileBuffer);
		CloseHandle(hFile);
	}
	return fSuccess;
}

// item information
BOOL CScriptHost::GetItemUnknown(
	LPCOLESTR		pItem,
	IUnknown	**	ppUnk)
{
	HRESULT			hr;
	if (NULL != this->m_pdispHost)
	{
		hr = this->m_pdispHost->QueryInterface(IID_IUnknown, (LPVOID*)ppUnk);
		return SUCCEEDED(hr);
	}
	else
	{
		*ppUnk = NULL;
		return FALSE;
	}
}

BOOL CScriptHost::GetItemTypeInfo(
	LPCOLESTR		pItem,
	ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	if (NULL != this->m_pdispHost)
	{
		hr = this->m_pdispHost->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, ppTypeInfo);
		return SUCCEEDED(hr);
	}
	else
	{
		*ppTypeInfo = NULL;
		return FALSE;
	}
}
