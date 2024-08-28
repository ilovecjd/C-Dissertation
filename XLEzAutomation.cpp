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


//Set cell value: Worksheet.Cells(nColumn, nRow).Value = szValue
BOOL CXLEzAutomation::SetCellValue(SheetName sheet, int nColumn, int nRow, CString szValue)
{
	return m_pXLServer->SetCellsValueToString(sheet, (double)nColumn, (double)nRow, szValue);
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
//Returns Worksheet.Cells(nColumn, nRow).Value
CString CXLEzAutomation::GetCellValue(SheetName sheet, int nColumn, int nRow)
{
	return m_pXLServer->GetCellValueCString(sheet, nColumn, nRow);
}
//Open Excell file
BOOL CXLEzAutomation::OpenExcelFile(CString szFileName)
{
	return m_pXLServer->OpenExcelFile(szFileName);
}
