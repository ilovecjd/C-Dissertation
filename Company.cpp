
#include "stdafx.h"
#include "C-Dissertation.h"
#include "XLEzAutomation.h"
#include "Company.h"
#include "Project.h"

CCompany::CCompany()
{
	// 동적 할당
	m_pGlobalEnv	= new GLOBAL_ENV;
	m_pXl			= new CXLEzAutomation;	
	m_pActType		= new ALL_ACT_TYPE;
	m_pActPattern	= new ALL_ACTIVITY_PATTERN;
	
}


CCompany::~CCompany()
{
	// 동적 할당된 메모리 해제
	delete m_pGlobalEnv;  // PGLOBAL_ENV 메모리 해제
	delete m_pXl;         // CXLEzAutomation 메모리 해제
	delete m_pActType;    // PACT_TYPE 메모리 해제
	delete m_pActPattern; // PALL_ACT_PATTERN 메모리 해제
	
}

// Proceed with the initialisation operation. This function shoule only be run once.
// shouldLoad is true, the function loads existing data.
// shouldLoad is false, the function creates new data.
//
// Parameters:
//   shouldLoad - A boolean flag indicating whether to load (true) or create (false).
//
// Returns:
//   A boolean value indicating the success (true) or failure (false) of the operation.

BOOL CCompany::Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad)
{	
	
	// song run once code 필요
	
	if (!m_pXl->OpenExcelFile(_T("d:\\1.xlsx")))
	{
		MessageBox(NULL, _T("Failed to open Excel file."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	if (m_pGlobalEnv == nullptr || pGlobalEnv == nullptr) {
		MessageBox(NULL, _T("pGlobalEnv is NULL."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}
	std::memcpy(m_pGlobalEnv, pGlobalEnv, sizeof(GLOBAL_ENV));
		
	m_pXl->ReadRangeToArray(ACTIVITY_STRUCT, 3, 2, (int*)m_pActType, 5, 13);
	m_pXl->ReadRangeToArray(ACTIVITY_STRUCT, 15, 2, (int*)m_pActPattern, 6, 26);
	
	CProject tempPrj;
	tempPrj.Init(m_pActType, m_pActPattern);
	// activitys 생성
	//CreateActivities
	return TRUE;
}
