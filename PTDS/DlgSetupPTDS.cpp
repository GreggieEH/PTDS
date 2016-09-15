#include "stdafx.h"
#include "DlgSetupPTDS.h"
#include "resource.h"
#include "dispids.h"

CDlgSetupPTDS::CDlgSetupPTDS(IUnknown * pUnk) :
	CMyDialog(IDD_DialogSetupPTDS),
	m_pUnk(NULL),
	m_dwPropNotify(0)
{
	if (NULL != pUnk)
	{
		this->m_pUnk = pUnk;
		this->m_pUnk->AddRef();
	}
}

CDlgSetupPTDS::~CDlgSetupPTDS()
{
	Utils_RELEASE_INTERFACE(this->m_pUnk);
}

BOOL CDlgSetupPTDS::OnInitDialog()
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	CImpIPropNotify*	pPropNotify;
	IUnknown	*	pUnkSink;

	Utils_CenterWindow(this->GetMyDialog());
	if (this->GetMyObject(&pdisp))
	{
		pPropNotify = new CImpIPropNotify(this);
		hr = pPropNotify->QueryInterface(IID_IPropertyNotifySink, (LPVOID*)&pUnkSink);
		if (SUCCEEDED(hr))
		{
			Utils_ConnectToConnectionPoint(pdisp, pUnkSink, IID_IPropertyNotifySink, TRUE, &this->m_dwPropNotify);
			pUnkSink->Release();
		}
		pdisp->Release();
	}
	// display values
	this->DisplayComment();
	this->DisplayUserName();
	this->DisplayStartWave();
	this->DisplayEndWave();
	this->DisplayStepSize();
	this->DisplayDwellTime();
	return TRUE;
}

BOOL CDlgSetupPTDS::OnCommand(
	WORD		wmID,
	WORD		wmEvent)
{
	switch (wmID)
	{
	case IDC_EDITCOMMENT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusComment();
			return TRUE;
		}
		break;
	case IDC_EDITUSERNAME:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusUserName();
			return TRUE;
		}
		break;
	case IDC_EDITSTARTWAVE:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusStartWave();
			return TRUE;
		}
		break;
	case IDC_EDITENDWAVE:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusEndWave();
			return TRUE;
		}
		break;
	case IDC_EDITSTEPSIZE:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusStepSize();
			return TRUE;
		}
		break;
	case IDC_EDITDWELLTIME:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->OnKillFocusDwellTime();
			return TRUE;
		}
		break;
	default:
		break;
	}
	return CMyDialog::OnCommand(wmID, wmEvent);
}

// EN_KILLFOCUS events
void CDlgSetupPTDS::OnKillFocusComment()
{
	IDispatch	*	pdisp;
	TCHAR			szComment[MAX_PATH];
	VARIANTARG		varg;

	if (this->ReadComment(szComment, MAX_PATH))
	{
		if (this->GetMyObject(&pdisp))
		{
			InitVariantFromString(szComment, &varg);
			Utils_InvokePropertyPut(pdisp, DISPID_Comment, &varg, 1);
			VariantClear(&varg);
			pdisp->Release();
		}
	}
	else
	{
		this->DisplayComment();
	}
}

void CDlgSetupPTDS::OnKillFocusUserName()
{
	IDispatch	*	pdisp;
	TCHAR			szUserName[MAX_PATH];
	VARIANTARG		varg;

	if (this->ReadUserName(szUserName, MAX_PATH))
	{
		if (this->GetMyObject(&pdisp))
		{
			InitVariantFromString(szUserName, &varg);
			Utils_InvokePropertyPut(pdisp, DISPID_UserName, &varg, 1);
			VariantClear(&varg);
			pdisp->Release();
		}
	}
	else
	{
		this->DisplayUserName();
	}
}

void CDlgSetupPTDS::OnKillFocusStartWave()
{
	double			startWave = this->ReadDoubleProperty(IDC_EDITSTARTWAVE);

	if (startWave >= 0.0)
	{
		this->SetDoubleProperty(DISPID_StartWave, startWave);
	}
	else
	{
		this->DisplayStartWave();
	}
}

void CDlgSetupPTDS::OnKillFocusEndWave()
{
	double			endWave = this->ReadDoubleProperty(IDC_EDITENDWAVE);

	if (endWave >= 0.0)
	{
		this->SetDoubleProperty(DISPID_EndWave, endWave);
	}
	else
	{
		this->DisplayEndWave();
	}
}

void CDlgSetupPTDS::OnKillFocusStepSize()
{
	double			stepSize = this->ReadDoubleProperty(IDC_EDITSTEPSIZE);
	if (stepSize > 0.0)
	{
		this->SetDoubleProperty(DISPID_StepSize, stepSize);
	}
	else
	{
		this->DisplayStepSize();
	}
}

void CDlgSetupPTDS::OnKillFocusDwellTime()
{
	double			dwellTime = this->ReadDoubleProperty(IDC_EDITDWELLTIME);
	this->SetDoubleProperty(DISPID_DwellTime, dwellTime);
}

void CDlgSetupPTDS::OnCloseup()
{
	IDispatch	*	pdisp;
	if (this->GetMyObject(&pdisp))
	{
		Utils_ConnectToConnectionPoint(pdisp, NULL, IID_IPropertyNotifySink, FALSE, &this->m_dwPropNotify);
		pdisp->Release();
	}
}

BOOL CDlgSetupPTDS::OnNotify(
	LPNMHDR		pnmh)
{
	return FALSE;
}

BOOL CDlgSetupPTDS::OnReturnClicked(
	UINT		nID)
{
	switch (nID)
	{
	case IDC_EDITCOMMENT:
		this->SetAllowClose(FALSE);
		SetFocus(GetDlgItem(this->GetMyDialog(), IDC_EDITUSERNAME));
		return TRUE;
	case IDC_EDITUSERNAME:
		this->SetAllowClose(FALSE);
		SetFocus(GetDlgItem(this->GetMyDialog(), IDC_EDITSTARTWAVE));
		return TRUE;
	case IDC_EDITSTARTWAVE:
		this->SetAllowClose(FALSE);
		SetFocus(GetDlgItem(this->GetMyDialog(), IDC_EDITENDWAVE));
		return TRUE;
	case IDC_EDITENDWAVE:
		this->SetAllowClose(FALSE);
		SetFocus(GetDlgItem(this->GetMyDialog(), IDC_EDITSTEPSIZE));
		return TRUE;
	case IDC_EDITSTEPSIZE:
		this->SetAllowClose(FALSE);
		SetFocus(GetDlgItem(this->GetMyDialog(), IDC_EDITDWELLTIME));
		return TRUE;
	case IDC_EDITDWELLTIME:
		this->OnKillFocusDwellTime();
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

// object access
BOOL CDlgSetupPTDS::GetMyObject(
	IDispatch**	ppdisp)
{
	HRESULT			hr;
	hr = this->m_pUnk->QueryInterface(IID_IDispatch, (LPVOID*)ppdisp);
	return SUCCEEDED(hr);
}

// on OK verifies the settings
BOOL CDlgSetupPTDS::OnOK()
{
	if (!CMyDialog::OnOK())
		return FALSE;
	TCHAR			szComment[MAX_PATH];
	TCHAR			szUserName[MAX_PATH];
	long			numSteps;
	TCHAR			szMessage[MAX_PATH];

	if (!this->GetComment(szComment, MAX_PATH))
	{
		MessageBox(this->GetMyDialog(), L"Comment not set", L"PTDS Setup", MB_OK);
		return FALSE;
	}
	if (!this->GetUserName(szUserName, MAX_PATH))
	{
		MessageBox(this->GetMyDialog(), L"User name not set", L"PTDS Setup", MB_OK);
		return FALSE;
	}
	numSteps = this->GetNumberOfSteps();
	if (numSteps <= 0)
	{
		StringCchPrintf(szMessage, MAX_PATH, L"Invalid number of steps = %1d", numSteps);
		MessageBox(this->GetMyDialog(), szMessage, L"PTDS Setup", MB_OK);
		return FALSE;
	}
	// survived all tests
	return TRUE;
}

// property change notification
void CDlgSetupPTDS::OnPropChanged(
	DISPID		dispid)
{
	switch (dispid)
	{
	case DISPID_StartWave:
		this->DisplayStartWave();
		break;
	case DISPID_EndWave:
		this->DisplayDwellTime();
		break;
	case DISPID_StepSize:
		this->DisplayStepSize();
		break;
	case DISPID_Comment:
		this->DisplayComment();
		break;
	case DISPID_DwellTime:
		this->DisplayDwellTime();
		break;
	case DISPID_UserName:
		this->DisplayUserName();
		break;
	default:
		break;
	}
}

// display properties
void CDlgSetupPTDS::DisplayComment()
{
	TCHAR			szComment[MAX_PATH];

	if (this->GetComment(szComment, MAX_PATH))
	{
		SetDlgItemText(this->GetMyDialog(), IDC_EDITCOMMENT, szComment);
	}
	else
	{
		SetDlgItemText(this->GetMyDialog(), IDC_EDITCOMMENT, L"");
	}
}

BOOL CDlgSetupPTDS::GetComment(
	LPTSTR		szComment,
	UINT		nBufferSize)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	size_t			slen;
	szComment[0] = '\0';
	if (this->GetMyObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_Comment, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToString(varResult, szComment, nBufferSize);
			if (SUCCEEDED(hr))
			{
				StringCchLength(szComment, nBufferSize, &slen);
				fSuccess = slen > 0;
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CDlgSetupPTDS::ReadComment(
	LPTSTR		szComment,
	UINT		nBufferSize)
{
	return (GetDlgItemText(this->GetMyDialog(), IDC_EDITCOMMENT, szComment, nBufferSize) > 0);
}

BOOL CDlgSetupPTDS::UpdateComment()
{
	TCHAR		szComment[MAX_PATH];
	IDispatch*	pdisp;
	BOOL		fSuccess = FALSE;
	VARIANTARG	varg;

	if (this->GetMyObject(&pdisp))
	{
		if (this->ReadComment(szComment, MAX_PATH))
		{
			InitVariantFromString(szComment, &varg);
			Utils_InvokePropertyPut(pdisp, DISPID_Comment, &varg, 1);
			VariantClear(&varg);
			fSuccess = TRUE;
		}
		pdisp->Release();
	}
	return fSuccess;
}

void CDlgSetupPTDS::DisplayUserName()
{
	TCHAR			szUserName[MAX_PATH];
	if (this->GetUserName(szUserName, MAX_PATH))
	{
		SetDlgItemText(this->GetMyDialog(), IDC_EDITUSERNAME, szUserName);
	}
	else
	{
		SetDlgItemText(this->GetMyDialog(), IDC_EDITUSERNAME, L"");
	}
}

BOOL CDlgSetupPTDS::GetUserName(
	LPTSTR		szUserName,
	UINT		nBufferSize)
{
	HRESULT			hr;
	IDispatch	*	pdisp;
	VARIANT			varResult;
	BOOL			fSuccess = FALSE;
	size_t			slen;
	szUserName[0] = '\0';
	if (this->GetMyObject(&pdisp))
	{
		hr = Utils_InvokePropertyGet(pdisp, DISPID_UserName, NULL, 0, &varResult);
		if (SUCCEEDED(hr))
		{
			hr = VariantToString(varResult, szUserName, nBufferSize);
			if (SUCCEEDED(hr))
			{
				StringCchLength(szUserName, nBufferSize, &slen);
				fSuccess = slen > 0;
			}
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
	return fSuccess;
}

BOOL CDlgSetupPTDS::ReadUserName(
	LPTSTR		szUserName,
	UINT		nBufferSize)
{
	return (GetDlgItemText(this->GetMyDialog(), IDC_EDITUSERNAME, szUserName, nBufferSize) > 0);
}


void CDlgSetupPTDS::DisplayStartWave()
{
	TCHAR			szString[MAX_PATH];
	double			dval = this->GetDoubleProperty(DISPID_StartWave);
	_stprintf_s(szString, L"%5.1f", dval);
	SetDlgItemText(this->GetMyDialog(), IDC_EDITSTARTWAVE, szString);
}

void CDlgSetupPTDS::DisplayEndWave()
{
	TCHAR			szString[MAX_PATH];
	double			dval = this->GetDoubleProperty(DISPID_EndWave);
	_stprintf_s(szString, L"%5.1f", dval);
	SetDlgItemText(this->GetMyDialog(), IDC_EDITENDWAVE, szString);
}

void CDlgSetupPTDS::DisplayStepSize()
{
	TCHAR			szString[MAX_PATH];
	double			dval = this->GetDoubleProperty(DISPID_StepSize);
	_stprintf_s(szString, L"%5.1f", dval);
	SetDlgItemText(this->GetMyDialog(), IDC_EDITSTEPSIZE, szString);
}

void CDlgSetupPTDS::DisplayDwellTime()
{
	TCHAR			szString[MAX_PATH];
	double			dval = this->GetDoubleProperty(DISPID_DwellTime);
	_stprintf_s(szString, L"%5.0f", dval);
	SetDlgItemText(this->GetMyDialog(), IDC_EDITDWELLTIME, szString);
}

double CDlgSetupPTDS::GetDoubleProperty(
	DISPID		dispid)
{
	IDispatch	*	pdisp;
	double			dval = -1.0;
	if (this->GetMyObject(&pdisp))
	{
		dval = Utils_GetDoubleProperty(pdisp, dispid);
		pdisp->Release();
	}
	return dval;
}

double CDlgSetupPTDS::ReadDoubleProperty(
	UINT		nID)
{
	TCHAR		szString[MAX_PATH];
	float		fval;
	double		dval = -1.0;

	if (GetDlgItemText(this->GetMyDialog(), nID, szString, MAX_PATH))
	{
		if (1 == _stscanf_s(szString, L"%f", &fval))
		{
			dval = (double)fval;
		}
	}
	return dval;
}

void CDlgSetupPTDS::SetDoubleProperty(
	DISPID		dispid,
	double		dval)
{
	IDispatch	*	pdisp;
	if (this->GetMyObject(&pdisp))
	{
		Utils_SetDoubleProperty(pdisp, dispid, dval);
		pdisp->Release();
	}
}

long CDlgSetupPTDS::GetNumberOfSteps()
{
	double		startWave = this->GetDoubleProperty(DISPID_StartWave);
	double		endWave = this->GetDoubleProperty(DISPID_EndWave);
	double		stepSize = this->GetDoubleProperty(DISPID_StepSize);
	double		numValues;

	if (0.0 == stepSize) return -1;
	numValues = ((endWave - startWave) / stepSize) + 1.0;
	return (long)floor(numValues);
}

CDlgSetupPTDS::CImpIPropNotify::CImpIPropNotify(CDlgSetupPTDS * pDlgSetupPTDS) :
	m_pDlgSetupPTDS(pDlgSetupPTDS),
	m_cRefs(0)
{

}

CDlgSetupPTDS::CImpIPropNotify::~CImpIPropNotify()
{

}

// IUnknown methods
STDMETHODIMP CDlgSetupPTDS::CImpIPropNotify::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyNotifySink == riid)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CDlgSetupPTDS::CImpIPropNotify::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CDlgSetupPTDS::CImpIPropNotify::Release()
{
	ULONG		cRefs = --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IPropertyNotifySink methods
STDMETHODIMP CDlgSetupPTDS::CImpIPropNotify::OnChanged(
	DISPID			dispID)
{
	this->m_pDlgSetupPTDS->OnPropChanged(dispID);
	return S_OK;
}


STDMETHODIMP CDlgSetupPTDS::CImpIPropNotify::OnRequestEdit(
	DISPID			dispID)
{
	return S_OK;
}