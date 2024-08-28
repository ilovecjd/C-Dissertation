// EzAutomation.h: interface for the CXLEzAutomation class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_EzAutomation_H__D140B9A3_1995_40AC_8E6D_8F23A95A63A2__INCLUDED_)
#define AFX_EzAutomation_H__D140B9A3_1995_40AC_8E6D_8F23A95A63A2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "XLAutomation.h"
#define xlNormal -4143

class CXLEzAutomation
{
public:
	BOOL OpenExcelFile(CString szFileName);
	CString GetCellValue(SheetName sheet, int nColumn, int nRow);
	BOOL SaveFileAs(CString szFileName);
	BOOL DeleteRow(SheetName sheet, int nRow);
	BOOL ReleaseExcel();
	BOOL SetCellValue(SheetName sheet, int nColumn, int nRow, CString szValue);
	BOOL ExportCString(SheetName sheet, CString szDataCollection);
	CXLEzAutomation();
	CXLEzAutomation(BOOL bVisible);
	BOOL ConvNumFormatColumn(int nColumn);
	virtual ~CXLEzAutomation();

protected:
	CXLAutomation* m_pXLServer;
};

#endif // !defined(AFX_EzAutomation_H__D140B9A3_1995_40AC_8E6D_8F23A95A63A2__INCLUDED_)
