#pragma once
#include "MyDialog.h"

class CDlgSetupPTDS :
	public CMyDialog
{
public:
	CDlgSetupPTDS(IUnknown * pUnk);
	virtual ~CDlgSetupPTDS();
protected:
	virtual BOOL		OnInitDialog();
	virtual BOOL		OnCommand(
							WORD		wmID,
							WORD		wmEvent);
	virtual BOOL		OnNotify(
							LPNMHDR		pnmh);
	virtual BOOL		OnReturnClicked(
							UINT		nID);
	// object access
	BOOL				GetMyObject(
							IDispatch**	ppdisp);
	// on OK verifies the settings
	virtual BOOL		OnOK();
	virtual void		OnCloseup();
	// property change notification
	void				OnPropChanged(
							DISPID		dispid);
	// display properties
	void				DisplayComment();
	BOOL				GetComment(
							LPTSTR		szComment,
							UINT		nBufferSize);
	BOOL				ReadComment(
							LPTSTR		szComment,
							UINT		nBufferSize);
	BOOL				UpdateComment();
	void				DisplayUserName();
	BOOL				GetUserName(
							LPTSTR		szUserName,
							UINT		nBufferSize);
	BOOL				ReadUserName(
							LPTSTR		szUserName,
							UINT		nBufferSize);
	void				DisplayStartWave();
	void				DisplayEndWave();
	void				DisplayStepSize();
	void				DisplayDwellTime();
	double				GetDoubleProperty(
							DISPID		dispid);
	double				ReadDoubleProperty(
							UINT		nID);
	void				SetDoubleProperty(
							DISPID		dispid,
							double		dval);
	long				GetNumberOfSteps();
	// EN_KILLFOCUS events
	void				OnKillFocusComment();
	void				OnKillFocusUserName();
	void				OnKillFocusStartWave();
	void				OnKillFocusEndWave();
	void				OnKillFocusStepSize();
	void				OnKillFocusDwellTime();


private:
	IUnknown		*	m_pUnk;
	DWORD				m_dwPropNotify;

	// property notifications
	class CImpIPropNotify : public IPropertyNotifySink
	{
	public:
		CImpIPropNotify(CDlgSetupPTDS * pDlgSetupPTDS);
		~CImpIPropNotify();
		// IUnknown methods
		STDMETHODIMP		QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		// IPropertyNotifySink methods
		STDMETHODIMP		OnChanged(
								DISPID			dispID);
		STDMETHODIMP		OnRequestEdit(
								DISPID			dispID);
	private:
		CDlgSetupPTDS	*	m_pDlgSetupPTDS;
		ULONG				m_cRefs;
	};
	friend CImpIPropNotify;
};

