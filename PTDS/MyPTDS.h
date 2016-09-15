#pragma once
#include "MyObject.h"

class CSciGratingScans;

class CMyPTDS : public CMyObject
{
public:
	CMyPTDS();
	virtual ~CMyPTDS();
	// initialize the object
	virtual HRESULT		Init();

	// sink events
	void				FireError(
							LPCTSTR		err);
protected:
	// hook to allow additional interfaces
	virtual HRESULT		_intQueryInterface(
							REFIID			riid,
							LPVOID		*	ppv);
	// invoker
	virtual HRESULT		InvokeHelper(
							DISPID		dispIdMember,
							WORD		wFlags,
							DISPPARAMS*	pDispParams,
							VARIANT	*	pVarResult);
	HRESULT				GetStartWave(
							VARIANT	*	pVarResult);
	HRESULT				SetStartWave(
							DISPPARAMS*	pDispParams);
	HRESULT				GetEndWave(
							VARIANT	*	pVarResult);
	HRESULT				SetEndWave(
							DISPPARAMS*	pDispParams);
	HRESULT				GetStepSize(
							VARIANT	*	pVarResult);
	HRESULT				SetStepSize(
							DISPPARAMS*	pDispParams);
	HRESULT				GetComment(
							VARIANT	*	pVarResult);
	HRESULT				SetComment(
							DISPPARAMS*	pDispParams);
	HRESULT				GetDwellTime(
							VARIANT	*	pVarResult);
	HRESULT				SetDwellTime(
							DISPPARAMS*	pDispParams);
	HRESULT				GetTimeConstant(
							VARIANT	*	pVarResult);
	HRESULT				SetTimeConstant(
							DISPPARAMS*	pDispParams);
	HRESULT				GetChopperFrequency(
							VARIANT	*	pVarResult);
	HRESULT				SetChopperFrequency(
							DISPPARAMS*	pDispParams);
	HRESULT				GetDataFile(
							VARIANT	*	pVarResult);
	HRESULT				SetDataFile(
							DISPPARAMS*	pDispParams);
	HRESULT				GetSlitWidth(
							VARIANT	*	pVarResult);
	HRESULT				SetSlitWidth(
							DISPPARAMS*	pDispParams);
	HRESULT				GetUserName(
							VARIANT	*	pVarResult);
	HRESULT				SetUserName(
							DISPPARAMS*	pDispParams);
	HRESULT				GetLampInfo(
							VARIANT	*	pVarResult);
	HRESULT				SetLampInfo(
							DISPPARAMS*	pDispParams);
	HRESULT				GetTimeStamp(
							VARIANT	*	pVarResult);
	HRESULT				SetTimeStamp(
							DISPPARAMS*	pDispParams);
	HRESULT				DisplaySetup(
							DISPPARAMS*	pDispParams,
							VARIANT	*	pVarResult);
	HRESULT				ScanStarted(
							DISPPARAMS*	pDispParams,
							VARIANT	*	pVarResult);
	HRESULT				AddDataValue(
							DISPPARAMS*	pDispParams,
							VARIANT	*	pVarResult);
	HRESULT				DisplayFilter(
							DISPPARAMS*	pDispParams);
	HRESULT				DisplayGrating(
							DISPPARAMS*	pDispParams);
	HRESULT				ScanEnded(
							DISPPARAMS*	pDispParams);
	HRESULT				WriteHeader(
							VARIANT	*	pVarResult);
	BOOL				WriteHeader(
							LPTSTR	*	szHeader);
	void				initObject(
							IDispatch*	pdispScript);
	BOOL				addProperty(
							IDispatch*	pdispScript,
							LPCTSTR		propName,
							VARIANT	*	propValue);
	BOOL				AddStringProperty(
							IDispatch*	pdispScript,
							LPCTSTR		szPropName,
							LPCTSTR		szPropValue);
	BOOL				AddDoubleProperty(
							IDispatch*	pdispScript,
							LPCTSTR		szPropName,
							double		dval);
	BOOL				AddTimeStamp(
							IDispatch*	pdispScript);
	BOOL				doWrite(
							IDispatch*	pdispScript,
							LPTSTR	*	szString);
	HRESULT				ReadHeader(
							DISPPARAMS*	pDispParams,
							VARIANT	*	pVarResult);
	BOOL				ReadHeader(
							LPCTSTR		szHeader);
	HRESULT				GetLockinInfo(
							VARIANT	*	pVarResult);
	HRESULT				SetLockinInfo(
							DISPPARAMS*	pDispParams);
	HRESULT				GetMonoInfo(
							VARIANT	*	pVarResult);
	HRESULT				SetMonoInfo(
							DISPPARAMS*	pDispParams);
	HRESULT				SetCurrentTime();
	BOOL				MakeFileName(
							LPTSTR		szFileName,
							UINT		nBufferSize);
	// append string to data file
	BOOL				WriteStringToDataFile(
							LPCTSTR			szString);
	// number resolution
	BOOL				CreateNumberResolution();
	void				SettestNumber(
							double			testNumber);
	BOOL				GetneedTenths();
	BOOL				GetneedHundreds();
	BOOL				GetneedThousands();
	double				scalePosition(
							double			position);
	// format a data value
	void				formatDataValue(
							double			wavelength,
							double			signal,
							LPTSTR			szValue,
							UINT			nBufferSize);
	// reform line ending
	void				ReformLineEnding(
							LPCTSTR			szInput,
							LPTSTR		*	szOutput);
	// write to data file
	BOOL				SaveToDataFile(
							LPCTSTR			szFilePath);
	// load from data file
	BOOL				LoadFromDataFile(
							LPCTSTR			szFilePath);
	HRESULT				ReadDataFile(
							DISPPARAMS	*	pDispParams,
							VARIANT		*	pVarResult);
	// browse header object from script
	BOOL				BrowseHeaderObject(
							IDispatch	*	pdisp);
	double				ReadDoubleProperty(
							IDispatch	*	pdisp,
							LPTSTR			szProperty);
	void				ReadStringProperty(
							IDispatch	*	pdisp,
							LPTSTR			szProperty,
							LPTSTR			szValue,
							UINT			nBufferSize);
	BOOL				ReadTimeStamp(
							IDispatch	*	pdisp);
private:
	// persistance interface
	class CImpIPersistFile : public IPersistFile
	{
	public:
		CImpIPersistFile(CMyPTDS * pMyPTDS);
		~CImpIPersistFile();
		// IUnknown methods
		STDMETHODIMP	QueryInterface(
							REFIID			riid,
							LPVOID		*	ppv);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		// IPersist method
		STDMETHODIMP	GetClassID(
							CLSID		*	pClassID);
		// IPersistFile methods
		STDMETHODIMP	GetCurFile(
							LPOLESTR	*	ppszFileName);
		STDMETHODIMP	IsDirty();
		STDMETHODIMP	Load(
							LPCOLESTR		pszFileName,
							DWORD			dwMode);
		STDMETHODIMP	Save(
							LPCOLESTR		pszFileName,
							BOOL			fRemember);
		STDMETHODIMP	SaveCompleted(
							LPCOLESTR		pszFileName);
	private:
		CMyPTDS		*	m_pMyPTDS;
		BOOL			m_fNoScribble;
	};
	friend CImpIPersistFile;
	CImpIPersistFile*	m_pImpIPersistFile;
	CSciGratingScans*	m_pSciGratingScans;			// grating scans object
	double				m_StartWave;
	double				m_EndWave;
	double				m_StepSize;
	TCHAR				m_szComment[MAX_PATH];
	double				m_DwellTime;
	double				m_TimeConstant;
	double				m_ChopperFrequency;
	TCHAR				m_szDataFile[MAX_PATH];
	double				m_SlitWidth;
	TCHAR				m_szUserName[MAX_PATH];
	TCHAR				m_szLampInfo[MAX_PATH];
	DATE				m_TimeStamp;
	TCHAR				m_szLockinInfo[MAX_PATH];
	TCHAR				m_szMonoInfo[MAX_PATH];
	BOOL				m_fAmScanning;
	IDispatch		*	m_pdispNumberRes;				// number resolution utility
	BOOL				m_fDirty;						// dirty flag
};

