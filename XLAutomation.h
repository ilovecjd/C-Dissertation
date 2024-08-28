// XLAutomation.h: interface for the CXLAutomation class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_XLAUTOMATION_H__E020CE95_7428_4BEF_A24C_48CE9323C450__INCLUDED_)
#define AFX_XLAUTOMATION_H__E020CE95_7428_4BEF_A24C_48CE9323C450__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// Sheet enumeration for easy reference
enum SheetName {
	PARAMETERS = 0,
	DASHBOARD,
	PROJECT,
	ACTIVITY_STRUCT,
	SHEET_COUNT // Total number of sheets
};

// CXLAutomation class definition
class CXLAutomation
{

#define MAX_DISP_ARGS 10
#define DISPARG_NOFREEVARIANT 0x01
#define DISP_FREEARGS 0x02
#define DISP_NOSHOWEXCEPTIONS 0x03
#define xlWorksheet -4167
#define xl3DPie -4102
#define xlRows 1
#define xlXYScatter -4169
#define xlXYScatterLines 74
#define xlXYScatterSmoothNoMarkers 73
#define xlXYScatterSmooth 72
#define xlXYScatterLinesNoMarkers 75
#define xlColumns 2
#define xlNormal -4143
#define xlUp -4162

#define xlContinuous 1 // xlContinuous for borders line style
#define xlThin 2       // xlThin for border weight
#define xlAutomatic -4105 // xlAutomatic for borders color

public:
	BOOL OpenExcelFile(CString szFileName);
	CString GetCellValueCString(SheetName sheet, int nColumn, int nRow);
	BOOL SaveAs(CString szFileName, int nFileFormat, CString szPassword, CString szWritePassword, BOOL bReadOnly, BOOL bBackUp);
	BOOL DeleteRow(SheetName sheet, long nRow);
	BOOL ReleaseExcel();
	BOOL UpdatePlotRange(SheetName sheet, int nYColumn);
	BOOL AddArgumentCStringArray(LPOLESTR lpszArgName, WORD wFlags, LPOLESTR *paszStrings, int iCount);
	BOOL SetRangeValueDouble(SheetName sheet, LPOLESTR lpszRef, double d);
	BOOL SetCellsValueToString(SheetName sheet, double Column, double Row, CString szStr);
	BOOL AddArgumentOLEString(LPOLESTR lpszArgName, WORD wFlags, LPOLESTR lpsz);
	BOOL AddArgumentCString(LPOLESTR lpszArgName, WORD wFlags, CString szStr);
	BOOL CreateWorkSheet();
	BOOL AddArgumentDouble(LPOLESTR lpszArgName, WORD wFlags, double d);
	BOOL AddArgumentBool(LPOLESTR lpszArgName, WORD wFlags, BOOL b);
	BOOL AddArgumentInt2(LPOLESTR lpszArgName, WORD wFlags, int i);
	BOOL AddArgumentDispatch(LPOLESTR lpszArgName, WORD wFlags, IDispatch* pdisp);
	void AddArgumentCommon(LPOLESTR lpszArgName, WORD wFlags, VARTYPE vt);
	BOOL InitOLE();

	//song 
	BOOL SetRangeValueAndStyle(SheetName sheet, int startRow, int startCol, int** dataArray, int numRows, int numCols);
	BOOL GetRange(SheetName sheet, int startRow, int startCol, int endRow, int endCol, VARIANTARG* pRange);
	//BOOL GetRange(SheetName sheet, CString rangeAddress, VARIANTARG* pRange);
	//BOOL ReadRangeToArray(SheetName sheet, int startRow, int startCol, int rowCount, int colCount, int** dataArray);
	//BOOL ReadRangeToArray(SheetName sheet, CString range, int* dataArray, int rows, int cols);
	BOOL ReadRangeToArray(SheetName sheet, int startRow, int startCol, int endRow, int endCol, int* dataArray, int rows, int cols);
	//song

	CXLAutomation();
	CXLAutomation(BOOL bVisible);
	virtual ~CXLAutomation();

protected:
	void ShowException(LPOLESTR szMember, HRESULT hr, EXCEPINFO *pexcep, unsigned int uiArgErr);
	void ReleaseDispatch();
	BOOL SetExcelVisible(BOOL bVisible);
	void ReleaseVariant(VARIANTARG *pvarg);
	void ClearAllArgs();
	void ClearVariant(VARIANTARG *pvarg);

	int m_iArgCount;
	int m_iNamedArgCount;
	VARIANTARG m_aVargs[MAX_DISP_ARGS];
	DISPID m_aDispIds[MAX_DISP_ARGS + 1]; // one extra for the member name
	LPOLESTR m_alpszArgNames[MAX_DISP_ARGS + 1]; // used to hold the argnames for GetIDs
	WORD m_awFlags[MAX_DISP_ARGS];

	BOOL ExlInvoke(IDispatch* pdisp, LPOLESTR szMember, VARIANTARG* pvargReturn,
		WORD wInvokeAction, WORD wFlags);
	IDispatch* m_pdispExcelApp;
	IDispatch* m_pdispWorkbook;
	IDispatch* m_pdispWorksheets[SHEET_COUNT]; // Array to store worksheet dispatch interfaces

	BOOL StartExcel();
	BOOL FindAndStoreWorksheet(IDispatch* pWorkbook, LPOLESTR sheetName, IDispatch** ppSheet);
};

#endif // !defined(AFX_XLAUTOMATION_H__E020CE95_7428_4BEF_A24C_48CE9323C450__INCLUDED_)
