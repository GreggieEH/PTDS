#pragma once
class CSciGratingScans
{
public:
	CSciGratingScans();
	~CSciGratingScans();
	BOOL			LoadFromText(
						LPCTSTR			szText);
	BOOL			SaveToText(
						LPTSTR		*	szText);
	BOOL			GetColumnData(
						long			scanIndex,
						long			column,
						ULONG		*	nData,
						double		**	ppData);
	long			GetnumberOfGratingScans();
	long			findScanIndex(
						long			grating,
						LPCTSTR			filter);
	long			getGratingNumber(
						long			scanIndex);
	BOOL			getFilter(
						long			scanIndex,
						LPTSTR			szFilter,
						UINT			nBufferSize);
	BOOL			addData(
						long			grating,
						LPCTSTR			filter,
						ULONG			nData,
						double		*	pData);
private:
	IDispatch	*	m_pdisp;
};

