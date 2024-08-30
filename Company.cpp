
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
	if (m_pXl != NULL)
		m_pXl->ReleaseExcel();
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

	CString strTitle[2][16] = {
		{
			_T("pType"), _T("PRJ_ID"), _T("기간"), _T("시작가능"), _T("끝"),
			_T("발주일"), _T("총수익"), _T("경험"), _T("성공%"), _T("CF갯수"),
			_T("CF1%"), _T("CF2%"), _T("CF3%"), _T("선금"), _T("중도"), _T("잔금")
		},
		{
			_T("act갯수"), _T(""), _T("Dur"), _T("start"), _T("end"),
			_T(""), _T("HR_H"), _T("HR_M"), _T("HR_L"), _T(""),
			_T("mon_cf1"), _T("mon_cf2"), _T("mon_cf3"), _T(""), _T(""), _T("")
		}
	};
	m_pXl->WriteArrayToRange(PROJECT, 1, 1, (CString*)strTitle, 2,16);

	//testFunction();
	PrintProjectInfo(&tempPrj);
	return TRUE;
}
void CCompany::PrintProjectInfo(CProject* pProject) {
	
	const int iWidth = 16;
	const int iHeight = 7;
	int posX, posY;

	VARIANT projectInfo[iHeight][iWidth];  // VARIANT 배열 생성
									 // 모든 VARIANT 요소를 VT_EMPTY로 초기화
	for (int i = 0; i < iHeight; ++i) {
		for (int j = 0; j < iWidth; ++j) {
			VariantInit(&projectInfo[i][j]);
			projectInfo[i][j].vt = VT_EMPTY;  // 초기화 상태를 VT_EMPTY로 설정
		}
	}

	// 첫 번째 행 설정	
	posX = 0; posY = 0;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->projectType;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->projectNum;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->projectDuration;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->possiblestartDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->endDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->orderDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = static_cast<int>(pProject->profit);
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->experience;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->successProbability;

	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->numCashFlows;									    
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_cashFlows[0];
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_cashFlows[1];
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_cashFlows[2];

	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->firstPayment;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->middlePayment;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->finalPayment;

	
	// 두 번째 행 설정
	posX = 0; posY = 1;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->numActivities;

	posX = 10;  // 빈 칸을 건너뛰기
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->firstPaymentMonth;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->middlePaymentMonth;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->finalPaymentMonth;
	
	// 활동 데이터 설정
	for (int index = 0; index < pProject->numActivities; ++index) {
		posX = 2; // 두 번째 행의 시작 위치
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[index].duration;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[index].startDate;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[index].endDate;

		posX = 6;  // 두 열 건너뛰기
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[index].highSkill;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[index].midSkill;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[index].lowSkill;

		posY++;
	}

	m_pXl->WriteArrayToRange(PROJECT, 4, 1, (VARIANT*)projectInfo, 7, 16);
}


void CCompany::testFunction()
{
	const int variantCount = 20;  // 주요 VARIANT 타입의 개수를 정의
	VARIANT variants[variantCount];  // VARIANT 배열 생성

									 // 모든 VARIANT 요소를 VT_EMPTY로 초기화
	for (int i = 0; i < variantCount; ++i) {
		VariantInit(&variants[i]);
		variants[i].vt = VT_EMPTY;  // 초기화 상태를 VT_EMPTY로 설정
	}

	// VARIANT 타입에 맞게 값 설정
	int index = 0;

	// VT_INT
	variants[index].vt = VT_INT;
	variants[index++].intVal = 42;

	index++;
	index++;

	// VT_I4 (32-bit signed integer)
	variants[index].vt = VT_I4;
	variants[index++].lVal = 100;

	// VT_R8 (64-bit floating-point number)
	variants[index].vt = VT_R8;
	variants[index++].dblVal = 3.14;

	// VT_BOOL (Boolean value)
	variants[index].vt = VT_BOOL;
	variants[index++].boolVal = VARIANT_TRUE;  // VARIANT_TRUE 또는 VARIANT_FALSE

											   // VT_BSTR (String)
	variants[index].vt = VT_BSTR;
	variants[index++].bstrVal = SysAllocString(L"Hello, VARIANT!");

	// VT_UI1 (8-bit unsigned integer)
	variants[index].vt = VT_UI1;
	variants[index++].bVal = 255;

	// VT_I2 (16-bit signed integer)
	variants[index].vt = VT_I2;
	variants[index++].iVal = 32767;

	// VT_UI2 (16-bit unsigned integer)
	variants[index].vt = VT_UI2;
	variants[index++].uiVal = 65535;

	// VT_UI4 (32-bit unsigned integer)
	variants[index].vt = VT_UI4;
	variants[index++].ulVal = 4294967295;

	// VT_I8 (64-bit signed integer)
	variants[index].vt = VT_I8;
	variants[index++].llVal = 9223372036854775807LL;

	// VT_UI8 (64-bit unsigned integer)
	variants[index].vt = VT_UI8;
	variants[index++].ullVal = 18446744073709551615ULL;

	// VT_R4 (32-bit floating-point number)
	variants[index].vt = VT_R4;
	variants[index++].fltVal = 2.71f;

	// VT_DATE (Date)
	variants[index].vt = VT_DATE;
	variants[index++].date = 44191.0;  // 예: 2021-01-01의 OLE 날짜

									   // VT_CY (Currency)
	variants[index].vt = VT_CY;
	variants[index++].cyVal.int64 = 10000;


	m_pXl->WriteArrayToRange(PROJECT, 1, 1, variants, 1, 20);

	// 메모리 정리
	for (int i = 0; i < variantCount; ++i) {
		VariantClear(&variants[i]);
	}

}
