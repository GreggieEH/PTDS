#include "stdafx.h"
#include "SciGratingScans.h"

CSciGratingScans::CSciGratingScans() :
	m_pdisp(NULL)
{
	HRESULT			hr;
	CLSID			clsid;

	hr = CLSIDFromProgID(L"Sciencetech.SciGratingScans.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&this->m_pdisp);
}

CSciGratingScans::~CSciGratingScans()
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

BOOL CSciGratingScans::LoadFromText(
	LPCTSTR			szText)
{
	HRESULT			hr;
	IPersistStreamInit*	pPersistStream;
	size_t			slen;
	IStream		*	pStream = NULL;
	BOOL			fSuccess = FALSE;

	if (NULL == this->m_pdisp) return FALSE;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*)&pPersistStream);
	if (SUCCEEDED(hr))
	{
		StringCbLength(szText, STRSAFE_MAX_CCH * sizeof(TCHAR), &slen);
		pStream = SHCreateMemStream((const BYTE*)szText, slen);
		if (NULL != pStream)
		{
			hr = pPersistStream->Load(pStream);
			fSuccess = SUCCEEDED(hr);
			pStream->Release();
		}
		pPersistStream->Release();
	}
	return fSuccess;
}

BOOL CSciGratingScans::SaveToText(
	LPTSTR		*	szText)
{
	HRESULT			hr;
	IPersistStreamInit*	pPersistStream;
	IStream		*	pStream = NULL;
	ULARGE_INTEGER	cbSize;
	BOOL			fSuccess = FALSE;
	STATSTG			statStg;
	DWORD			dwSize;
	DWORD			dwNRead;

	*szText = NULL;
	if (NULL == this->m_pdisp) return FALSE;
	hr = this->m_pdisp->QueryInterface(IID_IPersistStreamInit, (LPVOID*)&pPersistStream);
	if (SUCCEEDED(hr))
	{
		pPersistStream->GetSizeMax(&cbSize);
		pStream = SHCreateMemStream(NULL, 0);
		if (NULL != pStream)
		{
			pStream->SetSize(cbSize);
			hr = pPersistStream->Save(pStream, TRUE);
			if (SUCCEEDED(hr))
			{
				hr = pStream->Stat(&statStg, STATFLAG_NONAME);
				if (SUCCEEDED(hr))
				{
					dwSize = (DWORD)statStg.cbSize.QuadPart;
					*szText = (LPTSTR)CoTaskMemAlloc(dwSize + sizeof(TCHAR));
					ZeroMemory((PVOID)*szText, dwSize + sizeof(TCHAR));
					hr = pStream->Read((LPVOID)*szText, dwSize, &dwNRead);
					fSuccess = SUCCEEDED(hr);
				}
			}
			pStream->Release();
		}
		pPersistStream->Release();
	}
	return fSuccess;
}

BOOL CSciGratingScans::GetColumnData(
	long			scanIndex,
	long			column,
	ULONG		*	nData,
	double		**	ppData)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	*nData = 0;
	*ppData = NULL;
	Utils_GetMemid(this->m_pdisp, L"ColumnData", &dispid);
	InitVariantFromInt32(scanIndex, &avarg[1]);
	InitVariantFromInt32(column, &avarg[0]);
	hr = Utils_DoInvoke(this->m_pdisp, dispid, DISPATCH_PROPERTYGET, avarg, 2, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToDoubleArrayAlloc(varResult, ppData, nData);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

long CSciGratingScans::GetnumberOfGratingScans()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdisp, L"numberOfGratingScans", &dispid);
	return Utils_GetLongProperty(this->m_pdisp, dispid);
}

long CSciGratingScans::findScanIndex(
	long			grating,
	LPCTSTR			filter)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	long			scanIndex = -1;
	Utils_GetMemid(this->m_pdisp, L"findScanIndex", &dispid);
	InitVariantFromInt32(grating, &avarg[1]);
	InitVariantFromString(filter, &avarg[0]);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 2, &varResult);
	VariantClear(&avarg[0]);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &scanIndex);
	return scanIndex;
}

long CSciGratingScans::getGratingNumber(
	long			scanIndex)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	long			gratingNumber = -1;
	Utils_GetMemid(this->m_pdisp, L"getGratingNumber", &dispid);
	InitVariantFromInt32(scanIndex, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &gratingNumber);
	return gratingNumber;
}

BOOL CSciGratingScans::getFilter(
	long			scanIndex,
	LPTSTR			szFilter,
	UINT			nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	szFilter[0] = '\0';
	Utils_GetMemid(this->m_pdisp, L"getFilter", &dispid);
	InitVariantFromInt32(scanIndex, &varg);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szFilter, nBufferSize);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CSciGratingScans::addData(
	long			grating,
	LPCTSTR			filter,
	ULONG			nData,
	double		*	pData)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[3];
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	Utils_GetMemid(this->m_pdisp, L"addData", &dispid);
	InitVariantFromInt32(grating, &avarg[2]);
	InitVariantFromString(filter, &avarg[1]);
	InitVariantFromDoubleArray(pData, nData, &avarg[0]);
	hr = Utils_InvokeMethod(this->m_pdisp, dispid, avarg, 3, &varResult);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}
