#include "stdafx.h"
#include "MyPTDS.h"
#include "dispids.h"
#include "MyGuids.h"
#include "ScriptHost.h"
#include "SciGratingScans.h"

CMyPTDS::CMyPTDS() :
	CMyObject(),
	m_pImpIPersistFile(NULL),
	m_pSciGratingScans(NULL),		// grating scans object
	m_StartWave(0.0),
	m_EndWave(0.0),
	m_StepSize(1.0),
	m_DwellTime(0.0),
	m_TimeConstant(0.0),
	m_ChopperFrequency(0.0),
	m_SlitWidth(0.0),
	m_TimeStamp(0.0),
	m_fAmScanning(FALSE),
	m_pdispNumberRes(NULL)				// number resolution utility
{
	this->m_szComment[0] = '\0';
	this->m_szDataFile[0] = '\0';
	this->m_szUserName[0] = '\0';
	this->m_szLampInfo[0] = '\0';
	this->m_szLockinInfo[0] = '\0';
	this->m_szMonoInfo[0] = '\0';
}

CMyPTDS::~CMyPTDS()
{
	Utils_DELETE_POINTER(this->m_pImpIPersistFile);
	Utils_DELETE_POINTER(this->m_pSciGratingScans);
	Utils_RELEASE_INTERFACE(this->m_pdispNumberRes);
}

HRESULT CMyPTDS::Init()
{
	HRESULT			hr = CMyObject::Init();		// call the base class method
	if (SUCCEEDED(hr))
	{
		this->m_pImpIPersistFile = new CImpIPersistFile(this);
		this->m_pSciGratingScans = new CSciGratingScans();
	}
	return hr;
}

HRESULT CMyPTDS::_intQueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IPersist == riid || IID_IPersistFile == riid)
	{
		*ppv = this->m_pImpIPersistFile;
	}
	if (NULL != (*ppv))
	{
		((IUnknown*)(*ppv))->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

// sink events
void CMyPTDS::FireError(
	LPCTSTR		err)
{
	VARIANTARG			varg;
	InitVariantFromString(err, &varg);
	Utils_NotifySinks(this, MY_IIDSINK, DISPID_Error, &varg, 1);
	VariantClear(&varg);
}

HRESULT CMyPTDS::InvokeHelper(
	DISPID		dispIdMember,
	WORD		wFlags,
	DISPPARAMS*	pDispParams,
	VARIANT	*	pVarResult)
{
	switch (dispIdMember)
	{
	case DISPID_StartWave:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetStartWave(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetStartWave(pDispParams);
		break;
	case DISPID_EndWave:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetEndWave(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetEndWave(pDispParams);
		break;
	case DISPID_StepSize:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetStepSize(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetStepSize(pDispParams);
		break;
	case DISPID_Comment:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetComment(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetComment(pDispParams);
		break;
	case DISPID_DwellTime:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDwellTime(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDwellTime(pDispParams);
		break;
	case DISPID_TimeConstant:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetTimeConstant(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetTimeConstant(pDispParams);
		break;
	case DISPID_ChopperFrequency:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetChopperFrequency(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetChopperFrequency(pDispParams);
		break;
	case DISPID_DataFile:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetDataFile(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetDataFile(pDispParams);
		break;
	case DISPID_SlitWidth:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetSlitWidth(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetSlitWidth(pDispParams);
		break;
	case DISPID_UserName:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetUserName(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetUserName(pDispParams);
		break;
	case DISPID_LampInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetLampInfo(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetLampInfo(pDispParams);
		break;
	case DISPID_TimeStamp:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetTimeStamp(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetTimeStamp(pDispParams);
		break;
	case DISPID_DisplaySetup:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplaySetup(pDispParams, pVarResult);
		break;
	case DISPID_ScanStarted:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ScanStarted(pDispParams, pVarResult);
		break;
	case DISPID_AddDataValue:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->AddDataValue(pDispParams, pVarResult);
		break;
	case DISPID_DisplayFilter:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayGrating(pDispParams);
		break;
	case DISPID_DisplayGrating:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->DisplayGrating(pDispParams);
		break;
	case DISPID_ScanEnded:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ScanEnded(pDispParams);
		break;
	case DISPID_WriteHeader:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->WriteHeader(pVarResult);
		break;
	case DISPID_ReadHeader:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadHeader(pDispParams, pVarResult);
		break;
	case DISPID_LockinInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetLockinInfo(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetLockinInfo(pDispParams);
		break;
	case DISPID_MonoInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetMonoInfo(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetMonoInfo(pDispParams);
		break;
	case DISPID_SetCurrentTime:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SetCurrentTime();
		break;
	case DISPID_ReadDataFile:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->ReadDataFile(pDispParams, pVarResult);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyPTDS::GetStartWave(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_StartWave, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetStartWave(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_StartWave = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_StartWave);
	return S_OK;
}

HRESULT CMyPTDS::GetEndWave(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_EndWave, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetEndWave(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_EndWave = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_EndWave);
	return S_OK;
}

HRESULT CMyPTDS::GetStepSize(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_StepSize, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetStepSize(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_StepSize = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_StepSize);
	return S_OK;
}

HRESULT CMyPTDS::GetComment(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szComment, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetComment(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szComment, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this, DISPID_Comment);
	return S_OK;
}

HRESULT CMyPTDS::GetDwellTime(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_DwellTime, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetDwellTime(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_DwellTime = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_DwellTime);
	return S_OK;
}

HRESULT CMyPTDS::GetTimeConstant(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_TimeConstant, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetTimeConstant(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_TimeConstant = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_TimeConstant);
	return S_OK;
}

HRESULT CMyPTDS::GetChopperFrequency(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_ChopperFrequency, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetChopperFrequency(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_ChopperFrequency = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_ChopperFrequency);
	return S_OK;
}

HRESULT CMyPTDS::GetDataFile(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szDataFile, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetDataFile(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szDataFile, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this, DISPID_DataFile);
	return S_OK;
}

HRESULT CMyPTDS::GetSlitWidth(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_SlitWidth, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetSlitWidth(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_SlitWidth = varg.dblVal;
	Utils_OnPropChanged(this, DISPID_SlitWidth);
	return S_OK;
}

HRESULT CMyPTDS::GetUserName(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szUserName, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetUserName(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szUserName, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this, DISPID_UserName);
	return S_OK;
}

HRESULT CMyPTDS::GetLampInfo(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szLampInfo, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetLampInfo(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szLampInfo, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this, DISPID_LampInfo);
	return S_OK;
}

HRESULT CMyPTDS::GetTimeStamp(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(pVarResult);
	pVarResult->vt = VT_DATE;
	pVarResult->date = this->m_TimeStamp;
	return S_OK;
}

HRESULT CMyPTDS::SetTimeStamp(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_DATE, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_TimeStamp = varg.date;
	Utils_OnPropChanged(this, DISPID_TimeStamp);
	return S_OK;
}

HRESULT CMyPTDS::DisplaySetup(
	DISPPARAMS*	pDispParams,
	VARIANT	*	pVarResult)
{
	return S_OK;
}

HRESULT CMyPTDS::ScanStarted(
	DISPPARAMS*	pDispParams,
	VARIANT	*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	TCHAR			szFileName[MAX_PATH];		// file name generated by time stamp
	BOOL			fSuccess = FALSE;
	LPTSTR			szTemp = NULL;
	LPTSTR			szHeader = NULL;			// header

	// retrieve the working directory from the input
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	VariantToString(varg, this->m_szDataFile, MAX_PATH);
	VariantClear(&varg);
	// store the current time stamp
	this->SetCurrentTime();
	// generate the file name
	this->MakeFileName(szFileName, MAX_PATH);
	PathRenameExtension(szFileName, L".ptds");
	// full file path
	PathAppend(this->m_szDataFile, szFileName);
	// write the header info
	if (this->WriteHeader(&szTemp))
	{
		// replace \n with \r\n
		this->ReformLineEnding(szTemp, &szHeader);
		CoTaskMemFree((LPVOID)szTemp);
		// write to data file
		fSuccess = this->WriteStringToDataFile(szHeader);
		if (fSuccess)
		{
			this->CreateNumberResolution();
			this->SettestNumber(this->m_StepSize);
		}
		CoTaskMemFree((LPVOID)szHeader);
	}
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	if (fSuccess)
		this->m_fAmScanning = TRUE;
	return S_OK;
}

HRESULT CMyPTDS::AddDataValue(
	DISPPARAMS*	pDispParams,
	VARIANT	*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	double			xValue;
	double			signalValue;
	BOOL			fSuccess = FALSE;
	TCHAR			oneLine[MAX_PATH];
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	xValue = varg.dblVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	signalValue = varg.dblVal;
	this->formatDataValue(xValue, signalValue, oneLine, MAX_PATH);
	fSuccess = this->WriteStringToDataFile(oneLine);
	if (NULL != pVarResult)
		InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::DisplayFilter(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->WriteStringToDataFile(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyPTDS::DisplayGrating(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->WriteStringToDataFile(varg.bstrVal);
	VariantClear(&varg);
	return S_OK;
}

HRESULT CMyPTDS::ScanEnded(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->WriteStringToDataFile(varg.bstrVal);
	VariantClear(&varg);
	this->m_fAmScanning = FALSE;
	Utils_RELEASE_INTERFACE(this->m_pdispNumberRes);
	return S_OK;
}

HRESULT CMyPTDS::WriteHeader(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	LPTSTR			szHeader = NULL;
	if (this->WriteHeader(&szHeader))
	{
		InitVariantFromString(szHeader, pVarResult);
		CoTaskMemFree((LPVOID)szHeader);
	}
	return S_OK;
}

BOOL CMyPTDS::WriteHeader(
	LPTSTR	*	szHeader)
{
//	HRESULT			hr;
	BOOL			fSuccess = FALSE;
	CScriptHost		scriptHost(this);				// script host object
	TCHAR			szScriptFile[MAX_PATH];			// script file path
	IDispatch	*	pdispScript;					// script dispatch

	*szHeader = NULL;
	// form the script file name
	GetModuleFileName(GetOurInstance(), szScriptFile, MAX_PATH);
	PathRemoveFileSpec(szScriptFile);
	PathAppend(szScriptFile, L"headerInfo.jvs.txt");
	// load the script
	scriptHost.LoadScript(szScriptFile, NULL, NULL);
	if (scriptHost.GetScriptLoaded())
	{
		if (scriptHost.GetScriptDispatch(&pdispScript))
		{
			this->initObject(pdispScript);
			this->AddTimeStamp(pdispScript);
			this->AddDoubleProperty(pdispScript, L"TimeConstant", this->m_TimeConstant);
			this->AddStringProperty(pdispScript, L"Comment", this->m_szComment);
			this->AddStringProperty(pdispScript, L"UserName", this->m_szUserName);
			this->AddDoubleProperty(pdispScript, L"ChopperFrequency", this->m_ChopperFrequency);
			this->AddDoubleProperty(pdispScript, L"DwellTime", this->m_DwellTime);
			this->AddDoubleProperty(pdispScript, L"SlitWidtch", this->m_SlitWidth);
			this->AddStringProperty(pdispScript, L"LampInfo", this->m_szLampInfo);
			this->AddStringProperty(pdispScript, L"LockinInfo", this->m_szLockinInfo);
			this->AddStringProperty(pdispScript, L"MonoInfo", this->m_szMonoInfo);
			// write the string
			fSuccess = this->doWrite(pdispScript, szHeader);
			pdispScript->Release();
		}
	}
	return fSuccess;
}

HRESULT CMyPTDS::ReadHeader(
	DISPPARAMS*	pDispParams,
	VARIANT	*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fSuccess = this->ReadHeader(varg.bstrVal);
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

BOOL CMyPTDS::ReadHeader(
	LPCTSTR		szHeader)
{
	HRESULT			hr;
	BOOL			fSuccess = FALSE;
	CScriptHost		scriptHost(this);				// script host object
	TCHAR			szScriptFile[MAX_PATH];			// script file path
	IDispatch	*	pdispScript;					// script dispatch
	DISPID			dispid;
	VARIANT			varResult;
	VARIANTARG		varg;

	// form the script file name
	GetModuleFileName(GetOurInstance(), szScriptFile, MAX_PATH);
	PathRemoveFileSpec(szScriptFile);
	PathAppend(szScriptFile, L"headerInfo.jvs.txt");
	// load the script
	scriptHost.LoadScript(szScriptFile, NULL, NULL);
	if (scriptHost.GetScriptLoaded())
	{
		if (scriptHost.GetScriptDispatch(&pdispScript))
		{
			Utils_GetMemid(pdispScript, L"doRead", &dispid);
			InitVariantFromString(szHeader, &varg);
			Utils_InvokeMethod(pdispScript, dispid, &varg, 1, NULL);
			VariantClear(&varg);
			Utils_GetMemid(pdispScript, L"g_Obj", &dispid);
			hr = Utils_InvokePropertyGet(pdispScript, dispid, NULL, 0, &varResult);
			if (SUCCEEDED(hr))
			{
				if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
				{
					fSuccess = this->BrowseHeaderObject(varResult.pdispVal);
				}
				VariantClear(&varResult);
			}
			pdispScript->Release();
		}
	}
	return fSuccess;
}

// browse header object from script
BOOL CMyPTDS::BrowseHeaderObject(
	IDispatch	*	pdisp)
{

	this->m_TimeConstant = this->ReadDoubleProperty(pdisp, L"TimeConstant");
	this->ReadStringProperty(pdisp, L"Comment", this->m_szComment, MAX_PATH);
	this->ReadStringProperty(pdisp, L"UserName", this->m_szUserName, MAX_PATH);
	this->m_ChopperFrequency = this->ReadDoubleProperty(pdisp, L"ChopperFrequency");
	this->m_DwellTime = this->ReadDoubleProperty(pdisp, L"DwellTime");
	this->m_SlitWidth = this->ReadDoubleProperty(pdisp, L"SlitWidtch");
	this->ReadStringProperty(pdisp, L"LampInfo", this->m_szLampInfo, MAX_PATH);
	this->ReadStringProperty(pdisp, L"LockinInfo", this->m_szLockinInfo, MAX_PATH);
	this->ReadStringProperty(pdisp, L"MonoInfo", this->m_szMonoInfo, MAX_PATH);

	return TRUE;
}

double CMyPTDS::ReadDoubleProperty(
	IDispatch	*	pdisp,
	LPTSTR			szProperty)
{
	DISPID			dispid;
	Utils_GetMemid(pdisp, szProperty, &dispid);
	return Utils_GetDoubleProperty(pdisp, dispid);
}

void CMyPTDS::ReadStringProperty(
	IDispatch	*	pdisp,
	LPTSTR			szProperty,
	LPTSTR			szValue,
	UINT			nBufferSize)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	szValue[0] = '\0';
	Utils_GetMemid(pdisp, szProperty, &dispid);
	hr = Utils_InvokePropertyGet(pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		VariantToString(varResult, szValue, nBufferSize);
		VariantClear(&varResult);
	}
}

BOOL CMyPTDS::ReadTimeStamp(
	IDispatch	*	pdisp)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	TCHAR			szString[MAX_PATH];

	Utils_GetMemid(pdisp, L"TimeStamp", &dispid);
	hr = Utils_InvokePropertyGet(pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToString(varResult, szString, MAX_PATH);
		if (SUCCEEDED(hr))
		{
			hr = VarDateFromStr(szString, 0x1009, 0, &this->m_TimeStamp);
			fSuccess = SUCCEEDED(hr);
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

void CMyPTDS::initObject(
	IDispatch*	pdispScript)
{
	DISPID				dispid;
	Utils_GetMemid(pdispScript, L"initObject", &dispid);
	Utils_InvokeMethod(pdispScript, dispid, NULL, 0, NULL);
}


BOOL CMyPTDS::addProperty(
	IDispatch*	pdispScript,
	LPCTSTR		propName,
	VARIANT	*	propValue)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];

	Utils_GetMemid(pdispScript, L"addProperty", &dispid);
	InitVariantFromString(propName, &avarg[1]);
	VariantInit(&avarg[0]);
	VariantCopy(&avarg[0], propValue);
	hr = Utils_InvokeMethod(pdispScript, dispid, avarg, 2, NULL);
	VariantClear(&avarg[1]);
	VariantClear(&avarg[0]);
	return SUCCEEDED(hr);
}

BOOL CMyPTDS::AddStringProperty(
	IDispatch*	pdispScript,
	LPCTSTR		szPropName,
	LPCTSTR		szPropValue)
{
	BOOL		fSuccess = FALSE;
	VARIANT		Value;
	InitVariantFromString(szPropValue, &Value);
	fSuccess = this->addProperty(pdispScript, szPropName, &Value);
	VariantClear(&Value);
	return fSuccess;
}

BOOL CMyPTDS::AddDoubleProperty(
	IDispatch*	pdispScript,
	LPCTSTR		szPropName,
	double		dval)
{
	VARIANT		Value;
	InitVariantFromDouble(dval, &Value);
	return this->addProperty(pdispScript, szPropName, &Value);
}

BOOL CMyPTDS::AddTimeStamp(
	IDispatch*	pdispScript)
{
	HRESULT			hr;
	VARIANT			Value;
	BOOL			fSuccess = FALSE;

	VariantInit(&Value);
	Value.vt = VT_BSTR;
	hr = VarBstrFromDate(this->m_TimeStamp, 0x1009, 0, &Value.bstrVal);
	if (SUCCEEDED(hr))
	{
		fSuccess = this->addProperty(pdispScript, L"TimeStamp", &Value);
		VariantClear(&Value);
	}
	return fSuccess;
}

BOOL CMyPTDS::doWrite(
	IDispatch*	pdispScript,
	LPTSTR	*	szString)
{
	HRESULT		hr;
	DISPID		dispid;
	VARIANT		varResult;
	BOOL		fSuccess = FALSE;
	*szString = NULL;
	Utils_GetMemid(pdispScript, L"doWrite", &dispid);
	hr = Utils_InvokeMethod(pdispScript, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr))
	{
		hr = VariantToStringAlloc(varResult, szString);
		fSuccess = SUCCEEDED(hr);
		VariantClear(&varResult);
	}
	return fSuccess;
}

HRESULT CMyPTDS::GetLockinInfo(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szLockinInfo, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetLockinInfo(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szLockinInfo, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this, DISPID_LockinInfo);
	return S_OK;
}

HRESULT CMyPTDS::GetMonoInfo(
	VARIANT	*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromString(this->m_szMonoInfo, pVarResult);
	return S_OK;
}

HRESULT CMyPTDS::SetMonoInfo(
	DISPPARAMS*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(this->m_szMonoInfo, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	Utils_OnPropChanged(this, DISPID_MonoInfo);
	return S_OK;
}

HRESULT CMyPTDS::SetCurrentTime()
{
	SYSTEMTIME			st;

	GetLocalTime(&st);
	SystemTimeToVariantTime(&st, &this->m_TimeStamp);
	Utils_OnPropChanged(this, DISPID_TimeStamp);
	return S_OK;
}

// use the DateToString utility to make the file name
BOOL CMyPTDS::MakeFileName(
	LPTSTR		szFileName,
	UINT		nBufferSize)
{
	HRESULT			hr;
	CLSID			clsid;
	IDispatch	*	pdisp = NULL;
	DISPID			dispid;
	VARIANTARG		varg;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	szFileName[0] = '\0';
	hr = CLSIDFromProgID(L"Sci.DateToString.1", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&pdisp);
	if (SUCCEEDED(hr))
	{
		Utils_GetMemid(pdisp, L"dateToString", &dispid);
		VariantInit(&varg);
		varg.vt = VT_DATE;
		varg.date = this->m_TimeStamp;
		hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToString(varResult, szFileName, nBufferSize);
			fSuccess = SUCCEEDED(hr);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

// append string to data file
BOOL CMyPTDS::WriteStringToDataFile(
	LPCTSTR			szString)
{
	HANDLE			hFile;
	DWORD			dwPos;
	size_t			slen;
	LPTSTR			szText = NULL;
	DWORD			dwNWritten;
	BOOL			fSuccess = FALSE;

	hFile = CreateFile(this->m_szDataFile, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		// point to the end of the file
		dwPos = SetFilePointer(hFile, 0, NULL, FILE_END);

		// create the string to write
		StringCbLength(szString, STRSAFE_MAX_CCH * sizeof(TCHAR), &slen);
		szText = (LPTSTR)CoTaskMemAlloc(slen + 10);
		StringCbPrintf(szText, slen + 10, L"%s\r\n", szString);
		// write the string
		StringCbLength(szText, STRSAFE_MAX_CCH * sizeof(TCHAR), &slen);
		fSuccess = WriteFile(hFile, (LPCVOID)szText, slen, &dwNWritten, NULL);
		CoTaskMemFree((LPVOID)szText);
		CloseHandle(hFile);
	}
	return fSuccess;
}

// number resolution
BOOL CMyPTDS::CreateNumberResolution()
{
	HRESULT			hr;
	CLSID			clsid;

	Utils_RELEASE_INTERFACE(this->m_pdispNumberRes);
	hr = CLSIDFromProgID(L"Sciencetech.NumberRes.wsc.1.00", &clsid);
	if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*)&this->m_pdispNumberRes);
	return SUCCEEDED(hr);
}

void CMyPTDS::SettestNumber(
	double			testNumber)
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdispNumberRes, L"testNumber", &dispid);
	Utils_SetDoubleProperty(this->m_pdispNumberRes, dispid, testNumber);
}

BOOL CMyPTDS::GetneedTenths()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdispNumberRes, L"needTenths", &dispid);
	return Utils_GetBoolProperty(this->m_pdispNumberRes, dispid);
}

BOOL CMyPTDS::GetneedHundreds()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdispNumberRes, L"needHundreds", &dispid);
	return Utils_GetBoolProperty(this->m_pdispNumberRes, dispid);
}

BOOL CMyPTDS::GetneedThousands()
{
	DISPID			dispid;
	Utils_GetMemid(this->m_pdispNumberRes, L"needThousands", &dispid);
	return Utils_GetBoolProperty(this->m_pdispNumberRes, dispid);
}

double CMyPTDS::scalePosition(
	double			position)
{
	HRESULT			hr;
	DISPID			dispid;
	VARIANTARG		avarg[2];
	VARIANT			varResult;
	double			scaledPosition = position;
	Utils_GetMemid(this->m_pdispNumberRes, L"scalePosition", &dispid);
	InitVariantFromDouble(position, &avarg[1]);
	InitVariantFromDouble(this->m_StepSize, &avarg[0]);
	hr = Utils_InvokeMethod(this->m_pdispNumberRes, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToDouble(varResult, &scaledPosition);
	return scaledPosition;
}

// format a data value
void CMyPTDS::formatDataValue(
	double			wavelength,
	double			signal,
	LPTSTR			szValue,
	UINT			nBufferSize)
{
	TCHAR		szFormat[MAX_PATH];
	TCHAR		szWavelength[MAX_PATH];
	BOOL		fneedTenths = this->GetneedTenths();
	BOOL		fneedHundreds = this->GetneedHundreds();
	BOOL		fneedThousands = this->GetneedThousands();
	// display wavelength
	double		displayWave = this->scalePosition(wavelength);
	if (fneedThousands)
		StringCchCopy(szFormat, MAX_PATH, TEXT("   %6.3f"));
	else if (fneedHundreds)
		StringCchCopy(szFormat, MAX_PATH, TEXT("   %6.2f"));
	else
		StringCchCopy(szFormat, MAX_PATH, TEXT("   %6.1f"));
	_stprintf_s(szWavelength, MAX_PATH * sizeof(TCHAR), szFormat, displayWave);
	if (signal > 0.0)
	{
		_stprintf_s(szValue, nBufferSize, TEXT("%s    %-10.5e"), szWavelength, signal);
	}
	else
	{
		_stprintf_s(szValue, nBufferSize, TEXT("%s   %-10.5e"), szWavelength, signal);
	}
}

// reform line ending
void CMyPTDS::ReformLineEnding(
	LPCTSTR			szInput,
	LPTSTR		*	szOutput)
{
	LPTSTR		szScratch = NULL;			// scratch string
	size_t		slen;						// input string length
	size_t		i;
	size_t		t;

	*szOutput = NULL;
	StringCchLength(szInput, STRSAFE_MAX_CCH, &slen);
	// allocate lots of space for scratch
	szScratch = (LPTSTR)CoTaskMemAlloc(((slen * 2) + 1) * sizeof(TCHAR));
	ZeroMemory((PVOID)szScratch, ((slen * 2) + 1) * sizeof(TCHAR));
	for (i = 0, t = 0; i < slen && t < (slen * 2); i++)
	{
		if ('\n' == szInput[i])
		{
			szScratch[t] = '\r';
			t++;
			szScratch[t] = '\n';
			t++;
		}
		else
		{
			szScratch[t] = szInput[i];
			t++;
		}
	}
	StringCchLength(szScratch, STRSAFE_MAX_CCH, &slen);
	*szOutput = (LPTSTR)CoTaskMemAlloc((slen + 1) * sizeof(TCHAR));
	StringCchCopy(*szOutput, slen + 1, szScratch);
	CoTaskMemFree((LPVOID)szScratch);
}

// write to data file
BOOL CMyPTDS::SaveToDataFile(
	LPCTSTR			szFilePath)
{
	return FALSE;
}

// load from data file
BOOL CMyPTDS::LoadFromDataFile(
	LPCTSTR			szFilePath)
{
	BOOL			fSuccess = FALSE;
	LPTSTR			szFileBuffer = NULL;

	if (ReadFileIntoBuffer(szFilePath, &szFileBuffer))
	{
		// read the JSON section
		LPTSTR			szStart = StrChr(szFileBuffer, '{');
		LPTSTR			szEnd = StrChr(szFileBuffer, '}');
		LPTSTR			szRem = NULL;				// remainder
		if (NULL != szStart && NULL != szEnd)
		{
			size_t		slen1;
			size_t		slen2;
			size_t		nChars;
			LPTSTR		szHeader;
			StringCchLength(szStart, STRSAFE_MAX_CCH, &slen1);
			StringCchLength(szEnd, STRSAFE_MAX_CCH, &slen2);
			nChars = slen1 - slen2 + 2;
			szHeader = (LPTSTR)CoTaskMemAlloc((nChars + 1) * sizeof(TCHAR));
			StringCchCopyN(szHeader, nChars + 1, szStart, nChars);
			fSuccess = this->ReadHeader(szHeader);
			CoTaskMemFree((LPVOID)szHeader);
			szRem = &(szEnd[1]);
		}
		// check if successfully read in the header
		if (fSuccess)
		{
			// read in the grating scans
			fSuccess = this->m_pSciGratingScans->LoadFromText(szRem);
		}
		CoTaskMemFree(szFileBuffer);
	}
	return fSuccess;
}

HRESULT CMyPTDS::ReadDataFile(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	IPersistFile*	pPersistFile;
	BOOL			fSuccess = FALSE;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hr = this->QueryInterface(IID_IPersistFile, (LPVOID*)&pPersistFile);
	if (SUCCEEDED(hr))
	{
		hr = pPersistFile->Load(varg.bstrVal, 0);
		fSuccess = SUCCEEDED(hr);
		pPersistFile->Release();
	}
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

CMyPTDS::CImpIPersistFile::CImpIPersistFile(CMyPTDS * pMyPTDS) :
	m_pMyPTDS(pMyPTDS),
	m_fNoScribble(FALSE)
{

}

CMyPTDS::CImpIPersistFile::~CImpIPersistFile()
{

}

// IUnknown methods
STDMETHODIMP CMyPTDS::CImpIPersistFile::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_pMyPTDS->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyPTDS::CImpIPersistFile::AddRef()
{
	return this->m_pMyPTDS->AddRef();
}

STDMETHODIMP_(ULONG) CMyPTDS::CImpIPersistFile::Release()
{
	return this->m_pMyPTDS->Release();
}

// IPersist method
STDMETHODIMP CMyPTDS::CImpIPersistFile::GetClassID(
	CLSID		*	pClassID)
{
	*pClassID = MY_CLSID;
	return S_OK;
}

// IPersistFile methods
STDMETHODIMP CMyPTDS::CImpIPersistFile::GetCurFile(
	LPOLESTR	*	ppszFileName)
{
	size_t			slen;
	*ppszFileName = NULL;
	StringCchLength(this->m_pMyPTDS->m_szDataFile, MAX_PATH, &slen);
	if (slen > 0)
	{
		SHStrDup(this->m_pMyPTDS->m_szDataFile, ppszFileName);
		return S_OK;
	}
	else
	{
		SHStrDup(FILE_EXTENSION, ppszFileName);
		return S_FALSE;
	}
}

STDMETHODIMP CMyPTDS::CImpIPersistFile::IsDirty()
{
	return this->m_pMyPTDS->m_fDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CMyPTDS::CImpIPersistFile::Load(
	LPCOLESTR		pszFileName,
	DWORD			dwMode)
{
	HRESULT			hr = this->m_pMyPTDS->LoadFromDataFile(pszFileName);
	if (SUCCEEDED(hr))
	{
		StringCchCopy(this->m_pMyPTDS->m_szDataFile, MAX_PATH, pszFileName);
	}
	return hr;
}

STDMETHODIMP CMyPTDS::CImpIPersistFile::Save(
	LPCOLESTR		pszFileName,
	BOOL			fRemember)
{
	HRESULT			hr;
	if (NULL == pszFileName)
	{
		// save case
		if (this->m_fNoScribble) return E_UNEXPECTED;
		hr = this->m_pMyPTDS->SaveToDataFile(this->m_pMyPTDS->m_szDataFile) ? S_OK : E_FAIL;
		if (SUCCEEDED(hr)) this->m_fNoScribble = TRUE;
	}
	else
	{
		if (fRemember)
		{
			// save as
			hr = this->m_pMyPTDS->SaveToDataFile(pszFileName) ? S_OK : E_FAIL;
			if (SUCCEEDED(hr))
			{
				this->m_fNoScribble = TRUE;
				StringCchCopy(this->m_pMyPTDS->m_szDataFile, MAX_PATH, pszFileName);
			}
		}
		else
		{
			// save a copy as
			hr = this->m_pMyPTDS->SaveToDataFile(pszFileName) ? S_OK : E_FAIL;
		}
	}
	return S_OK;
}

STDMETHODIMP CMyPTDS::CImpIPersistFile::SaveCompleted(
	LPCOLESTR		pszFileName)
{
	this->m_fNoScribble = FALSE;
	return S_OK;
}