// EzAutomation.cpp: implementation of the CXLEzAutomation class.
//This wrapper classe is provided for easy access to basic automation  
//methods of the CXLAutoimation.
//Only very basic set of methods is provided here.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "C-Dissertation.h"
#include "XLEzAutomation.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//

CXLEzAutomation::CXLEzAutomation()
{
	//Starts Excel with bVisible = TRUE and creates empty worksheet 
	m_pXLServer = new CXLAutomation;
}

CXLEzAutomation::CXLEzAutomation(BOOL bVisible)
{
	//Can be used to start Excel in background (bVisible = FALSE)
	m_pXLServer = new CXLAutomation(bVisible);

}

CXLEzAutomation::~CXLEzAutomation()
{
	if(NULL != m_pXLServer)
		delete m_pXLServer;
}



//Quit Excel
BOOL CXLEzAutomation::ReleaseExcel()
{
	return m_pXLServer->ReleaseExcel();
}
//Delete line from worksheet
BOOL CXLEzAutomation::DeleteRow(SheetName sheet, int nRow)
{
	return m_pXLServer->DeleteRow(sheet, nRow);
}
//Save workbook as Excel file
BOOL CXLEzAutomation::SaveFileAs(CString szFileName)
{
	return m_pXLServer->SaveAs(szFileName, xlNormal, _T(""), _T(""), FALSE, FALSE);
}

//Open Excell file
BOOL CXLEzAutomation::OpenExcelFile(CString szFileName)
{
	return m_pXLServer->OpenExcelFile(szFileName);
}

// Overloaded GetCellValue functions
// Returns integer value from Worksheet.Cells(nColumn, nRow)
BOOL CXLEzAutomation::GetCellValue(SheetName sheet, int nColumn, int nRow, int* pValue)
{
	if (pValue == nullptr) return FALSE;
	return m_pXLServer->GetCellValueInt(sheet, nColumn, nRow, pValue);
}

// Returns CString value from Worksheet.Cells(nColumn, nRow)
BOOL CXLEzAutomation::GetCellValue(SheetName sheet, int nColumn, int nRow, CString* pValue)
{
	if (pValue == nullptr) return FALSE;
	return m_pXLServer->GetCellValueCString(sheet, nColumn, nRow, pValue);
}

// Returns double value from Worksheet.Cells(nColumn, nRow)
BOOL CXLEzAutomation::GetCellValue(SheetName sheet, int nColumn, int nRow, double* pValue)
{
	if (pValue == nullptr) return FALSE;
	return m_pXLServer->GetCellValueDouble(sheet, nColumn, nRow, pValue);
}


// SetCellValue for integer
BOOL CXLEzAutomation::SetCellValue(SheetName sheet, int nColumn, int nRow, int value)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->SetCellValueInt(sheet, nColumn, nRow, value);
}

// SetCellValue for CString
BOOL CXLEzAutomation::SetCellValue(SheetName sheet, int nColumn, int nRow, CString value)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->SetCellValueCString(sheet, nColumn, nRow, value);
}

// SetCellValue for double
BOOL CXLEzAutomation::SetCellValue(SheetName sheet, int nColumn, int nRow, double value)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->SetCellValueDouble(sheet, nColumn, nRow, value);
}

// Overloaded function to read integer values from Excel
BOOL CXLEzAutomation::ReadRangeToArray(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols)
{
	if (!m_pXLServer) return FALSE;
	return m_pXLServer->ReadRangeToIntArray(sheet, startRow, startCol, dataArray, rows, cols);
}

// Overloaded function to read CString values from Excel
BOOL CXLEzAutomation::ReadRangeToArray(SheetName sheet, int startRow, int startCol, CString* dataArray, int rows, int cols)
{
	if (!m_pXLServer) return FALSE;
	return m_pXLServer->ReadRangeToCStringArray(sheet, startRow, startCol, dataArray, rows, cols);
}

// int 배열을 Excel에 쓰기
BOOL CXLEzAutomation::WriteArrayToRange(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->WriteArrayToRangeInt(sheet, startRow, startCol, dataArray, rows, cols);
}

// CString 배열을 Excel에 쓰기
BOOL CXLEzAutomation::WriteArrayToRange(SheetName sheet, int startRow, int startCol, CString* dataArray, int rows, int cols)
{
	if (m_pXLServer == NULL)
		return FALSE;
	return m_pXLServer->WriteArrayToRangeCString(sheet, startRow, startCol, dataArray, rows, cols);
}
