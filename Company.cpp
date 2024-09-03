
#include "stdafx.h"
#include "C-Dissertation.h"
#include "GlobalEnv.h"
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
	//DeallocateManageTable(&m_manageTable); //song 확인하고 지우자

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
	/////////////////////////////////////////////////////////////////////////
	// 전달 받은 환경 변수를 Company 로 복사
	*m_pGlobalEnv = *pGlobalEnv;		
	m_pXl->ReadRangeToArray(WS_NUM_ACTIVITY_STRUCT, 3, 2, (int*)m_pActType, 5, 13);
	m_pXl->ReadRangeToArray(WS_NUM_ACTIVITY_STRUCT, 15, 2, (int*)m_pActPattern, 6, 26);

	AllTableInit(m_pGlobalEnv->SimulationWeeks); //dash boar 용 배열들의 크기 조절	

	if (shouldLoad)
		LoadProjectsFromExcel();
	else
		CreateProjects();
	return TRUE;
}

void CCompany::PrintProjectInfo(SheetName sheet, CProject* pProject) {
	
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
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_category;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_ID;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_duration;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_startAvail;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_endDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_orderDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = static_cast<int>(pProject->m_profit);
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_experience;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_winProb;

	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_nCashFlows;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_cashFlows[0];
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_cashFlows[1];
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_cashFlows[2];

	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_firstPay;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_secondPay;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_finalPay;

	
	// 두 번째 행 설정
	posX = 0; posY = 1;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->numActivities;

	posX = 10;  // 빈 칸을 건너뛰기
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_firstPayMonth;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_secondPayMonth;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_finalPayMonth;
	
	posX = 14;  // 빈 칸을 건너뛰기
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_projectType;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activityPattern;
	
	// 활동 데이터 설정
	for (int i = 0; i < pProject->numActivities; ++i) {
		// 인덱스를 문자열로 변환하고 "Activity" 접두사 추가
		CString strAct;		
		strAct.Format(_T("Activity%02d"), i + 1);

		posX = 1; // 엑셀의 2행 2열부터 적는다.
		projectInfo[posY][posX].vt = VT_BSTR; projectInfo[posY][posX++].bstrVal = strAct.AllocSysString();
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[i].duration;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[i].startDate;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[i].endDate;

		posX = 6;  // 두 열 건너뛰기
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[i].highSkill;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[i].midSkill;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->m_activities[i].lowSkill;

		posY++;
	}

	int printY = 4 + (pProject->m_ID -1)*iHeight;
	m_pXl->WriteArrayToRange(sheet, printY, 1, (VARIANT*)projectInfo, iHeight, iWidth);
	m_pXl->SetRangeBorderAround(sheet, printY, 1, printY + iHeight-1, iWidth + 1 - 1, 1, 2, RGB(0, 0, 0));
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


	m_pXl->WriteArrayToRange(WS_NUM_PROJECT, 1, 1, variants, 1, 20);

	// 메모리 정리
	for (int i = 0; i < variantCount; ++i) {
		VariantClear(&variants[i]);
	}

}



//// Function to dynamically allocate memory for all int* members of the struct
//void CCompany:: AllocateManageTable(MANAGE_TABLE* table, int size) {
//	// Calculate the number of int* members dynamically
//	int memberCount = sizeof(MANAGE_TABLE) / sizeof(int*);
//
//	// Pointer to the start of the struct
//	char* baseAddress = reinterpret_cast<char*>(table);
//
//	// Loop through each int* member and allocate memory
//	for (int i = 0; i < memberCount; ++i) {
//		int** memberPtr = reinterpret_cast<int**>(baseAddress + i * sizeof(int*));
//		*memberPtr = new int[size]; // Allocate memory for each member
//	}
//}
//
//// Function to deallocate memory for all int* members of the struct
//void CCompany::DeallocateManageTable(MANAGE_TABLE* table) 
//{
//	// Calculate the number of int* members dynamically
//	int memberCount = sizeof(MANAGE_TABLE) / sizeof(int*);
//
//	// Pointer to the start of the struct
//	char* baseAddress = reinterpret_cast<char*>(table);
//
//	// Loop through each int* member and deallocate memory
//	for (int i = 0; i < memberCount; ++i) 
//	{
//		int** memberPtr = reinterpret_cast<int**>(baseAddress + i * sizeof(int*));
//		delete[] * memberPtr; // Deallocate memory for each member
//		*memberPtr = nullptr; // Set pointer to nullptr to avoid dangling pointer
//	}
//}

BOOL CCompany::CreateProjects()
{
	int cnt = 0, sum = 0;
	int lastWeek = m_pGlobalEnv->SimulationWeeks;

	//AllocateManageTable(&m_manageTable, lastWeek);

	/////////////////////////////////////////////////////////////////////////
	// 프로젝트 발주(발생) 현황 생성
	for (int week = 0; week < lastWeek; week++)
	{
		cnt = PoissonRandom(m_pGlobalEnv->WeeklyProb); //        ' 이번주 발생하는 프로젝트 갯수		
		m_orderTable[ORDER_SUM][week] = sum;	//' 누계
		m_orderTable[ORDER_ORD][week] = cnt;	//' 발생 프로젝트갯수

											//' 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다. ==> 이전주까지 발생한 프로젝트 갯수후위연산. vba에서 do while 문법 모름... ㅎㅎ
		sum = sum + cnt;
	}
	m_totalProjectNum = sum;


	PrintDBTitle();

	
	/////////////////////////////////////////////////////////////////////////
	// project 시트에 헤더 출력	
	CString strTitle[2][16] = {
		{
			_T("Category"), _T("PRJ_ID"), _T("기간"), _T("시작가능"), _T("끝"),
			_T("발주일"), _T("총수익"), _T("경험"), _T("성공%"), _T("CF갯수"),
			_T("CF1%"), _T("CF2%"), _T("CF3%"), _T("선금"), _T("중도"), _T("잔금")
		},
		{
			_T("act갯수"), _T(""), _T("Dur"), _T("start"), _T("end"),
			_T(""), _T("HR_H"), _T("HR_M"), _T("HR_L"), _T(""),
			_T("mon_cf1"), _T("mon_cf2"), _T("mon_cf3"), _T(""), _T("prjType"), _T("actType")
		}
	};
	m_pXl->WriteArrayToRange(WS_NUM_PROJECT, 1, 1, (CString*)strTitle, 2, 16);
	m_pXl->SetRangeBorder(WS_NUM_PROJECT, 1, 1, 2, 16, 1, xlThin, RGB(0, 0, 0));


	/////////////////////////////////////////////////////////////////////////
	// 프로젝트 생성
	// song ==> null 체크 하자.
	m_AllProjects = new CProject*[sum];

	int projectId = 0;
	int startNum = 0;
	int endNum = 0;
	int preTotal = 0;

	for (int week = 0; week < lastWeek; week++)
	{
		preTotal = m_orderTable[ORDER_SUM][week];			// 지난주까지의 발주 프로젝트 누계
		startNum = preTotal + 1;						// 신규프로젝트이 시작번호 = 누계 +1
		endNum = preTotal + m_orderTable[ORDER_ORD][week];	// 마지막 프로젝트의 시작번호 = 지난주 누계 + 이번주 발생건수

		if ((startNum != 0) && (startNum <= endNum))
		{
			for (projectId = startNum; projectId <= endNum; projectId++)
			{
				CProject* pTempPrj;
				pTempPrj = new CProject;
				pTempPrj->Init(0, projectId, week, m_pActType, m_pActPattern);

				m_AllProjects[projectId - 1] = pTempPrj;
				PrintProjectInfo(WS_NUM_PROJECT, pTempPrj);
			}
		}
	}
	
	return TRUE;
}

// song
BOOL CCompany::LoadProjectsFromExcel()
{
	int cnt = 0, sum = 0;
	int lastWeek = m_pGlobalEnv->SimulationWeeks;
	
	/////////////////////////////////////////////////////////////////////////
	// 프로젝트 발주(발생) 현황 로드
	 //다음 내용을 가져오자
	//ReadRangeToArray(SheetName sheet, int startRow, int startCol, int* dataArray, int rows, int cols)
	int* tempBuf = new int[ORDER_COUNT*lastWeek];
	m_pXl->ReadRangeToArray(WS_NUM_DASHBOARD, 3, 2, tempBuf, 2, lastWeek);
	//m_pXl->ReadRangeToArray(WS_NUM_DASHBOARD, 2, 2, m_manageTable.pWeeksNum, 1, lastWeek);
	//m_pXl->ReadRangeToArray(WS_NUM_DASHBOARD, 3, 2, m_manageTable.pSum, 1, lastWeek);
	//m_pXl->ReadRangeToArray(WS_NUM_DASHBOARD, 4, 2, m_manageTable.pOrder, 1, lastWeek);
	m_orderTable.copyFromContinuousMemory(tempBuf, ORDER_COUNT, lastWeek);
	m_totalProjectNum = m_orderTable[ORDER_SUM][lastWeek-1] + m_orderTable[ORDER_ORD][lastWeek-1];
	
	/////////////////////////////////////////////////////////////////////////
	// project 시트에 헤더 출력
	
	/////////////////////////////////////////////////////////////////////////
	// 프로젝트 생성 후 내용은 로드
	// song ==> project 의 생성자와 소멸자, init 함수를 확인해 놓자.
	
	// song ==> NULL 체크 하자
	m_AllProjects = new CProject*[m_totalProjectNum];
	int* pProjectInfo;

	LONG lInfoSize = 7 * m_totalProjectNum * 16;
	pProjectInfo = new int[lInfoSize];
	m_pXl->ReadExRangeConvertInt(WS_NUM_PROJECT, 4, 1, pProjectInfo, m_totalProjectNum * 7, 16);

	for (int i = 0; i < m_totalProjectNum; i++)
	{	
		LONG lBaseAddress = 0;
		LONG lTemp = 0;
		
		lBaseAddress = 7* i * 16;

		CProject* pTempPrj;
		pTempPrj = new CProject;
		//pTempPrj->Init( m_pActType, m_pActPattern);

		// 첫 번째 행 설정			
		pTempPrj->m_category				= *(pProjectInfo+lBaseAddress++);
		pTempPrj->m_ID						= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_duration				= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_startAvail				= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_endDate					= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_orderDate				= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_profit					= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_experience	= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_winProb		= *(pProjectInfo + lBaseAddress++);
		
		pTempPrj->m_nCashFlows	= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_cashFlows[0]	= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_cashFlows[1]	= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_cashFlows[2]	= *(pProjectInfo + lBaseAddress++);
		
		pTempPrj->m_firstPay		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_secondPay		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->m_finalPay		= *(pProjectInfo + lBaseAddress++);

		// 두 번째 행 
		lTemp = lBaseAddress;

		pTempPrj->numActivities = *(pProjectInfo + lBaseAddress++);
		
		// 활동 데이터 설정
		for (int j = 0; j < pTempPrj->numActivities; j++)
		{
			lBaseAddress += 1;// 빈 칸을 건너뛰기 (Activity01)
			pTempPrj->m_activities[j].duration = *(pProjectInfo + lBaseAddress++);
			pTempPrj->m_activities[j].startDate = *(pProjectInfo + lBaseAddress++);
			pTempPrj->m_activities[j].endDate = *(pProjectInfo + lBaseAddress++);
			
			lBaseAddress += 1;  // 빈 칸을 건너뛰기
			pTempPrj->m_activities[j].highSkill = *(pProjectInfo + lBaseAddress++);
			pTempPrj->m_activities[j].midSkill = *(pProjectInfo + lBaseAddress++);
			pTempPrj->m_activities[j].lowSkill = *(pProjectInfo + lBaseAddress++);

			if (j == 0)
			{
				lBaseAddress += 1;  // 빈 칸을 건너뛰기
				pTempPrj->m_firstPayMonth = *(pProjectInfo + lBaseAddress++);
				pTempPrj->m_secondPayMonth = *(pProjectInfo + lBaseAddress++);
				pTempPrj->m_finalPayMonth = *(pProjectInfo + lBaseAddress++);

				lBaseAddress += 1;  // 빈 칸을 건너뛰기
				pTempPrj->m_projectType = *(pProjectInfo + lBaseAddress++);
				pTempPrj->m_activityPattern = *(pProjectInfo + lBaseAddress++);

				lBaseAddress += 1;  // 빈 칸을 건너뛰기
			}
			else 
			{
				lBaseAddress += 8;  // 빈 칸을 건너뛰기
			}
		}

		pTempPrj->m_isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
		
		m_AllProjects[i] = pTempPrj;
		PrintProjectInfo(WS_NUM_DEBUG_INFO, pTempPrj);
		lBaseAddress = lTemp + 6 * 16;
	}
	
	delete[] pProjectInfo;
	pProjectInfo = NULL;

	return TRUE;
}

// 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
void CCompany::Decision(int thisWeek ) {

	if (0 == thisWeek) // 첫주는 체크할 지난주가 없음
		return;

	// 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
	CheckLastWeek(thisWeek);

	// 2. 진행 가능한 후보프로젝트들을  candidateTable에 모은다
	//SelectionOfCandidates(thisWeek);

	// 3. 신규 프로젝트 선택 및 진행프로젝트 업데이트
	// 이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
	//SelectNewProject(thisWeek);

	// Call comPrintDashboard()
	
}

// 완료프로젝트 검사 및 진행프로젝트 업데이트
// 1. 지난 기간의 정보를 이번기간에 복사하고
// 2. 지난 기간에 진행중인 프로젝트중 완료된 것이 있는가?
// 3. 완료된 프로젝트들만 이번기간에서 삭제
void CCompany::CheckLastWeek(int thisWeek )
{	
	// 수입과 지출 테이블은 매주 업데이트 한다.
	//m_incomeTable(thisWeek) = m_totalIncome;
	//m_costsTable(thisWeek) = m_Totalcosts;

	
	int nLastProjects = m_doingTable[0][thisWeek - 1];//지난주에 진행 중이던 프로젝트의 갯수
	if (0 == nLastProjects)//song ==> 지난주에 진행중이던 프로젝트가 없다.
		return;

	for (int i = 0; i < nLastProjects; i++)
	{
		int prjId = m_doingTable[i + 1][thisWeek - 1];
		if (prjId == 0)
			return;

		CProject* project = m_AllProjects[prjId];

		// 1. payment 를 계산한다. 선금은 시작시 받기로 한다. 조건완료후 1주 후 수금			
		// 2. 지출을 계산한다.
		//' 3. 진행중인 프로젝트를 이관해서 기록한다.
		if (thisWeek < (project->m_isStart + project->m_duration - 1)) // ' 아직 안끝났으면
		{
			int sum = m_doingTable[0][thisWeek];
			m_doingTable[sum][thisWeek] = project->m_ID;// 테이블 크기는 자동으로 변경된다.
			sum += 1;
			m_doingTable[0][thisWeek] = sum;
		}
	}
}


//
//void CCompany::SelectCandidates(thisWeek)
//{
//	for (int i = 0 ; i< MAX_CANDIDATES; i++)
//	{
//		m_candidateTable[i] = 0;
//	}
//
//	int sum = m_manageTable.pSum[week];	// 누계
//	int nOrder = m_manageTable.pOrder[week];	//' 발생 프로젝트갯수
//
//	startProjectNum = sum + 1;  // 이번기간의 처음 프로젝트
//	endProjectNum = m_manageTable.pOrder[week] + cnt;  // 이번기간의 마지막 프로젝트
//
//	for (int id = startProjectNum; id < endProjectNum; id++)
//	{
//		project = m_projectTable[id - 1];
//
//		if () // 인원 체크
//		{
//
//		}
//	}
//	
//
//}


// dash boar 용 배열들의 크기 조절	
void CCompany::AllTableInit(int nWeeks)
{
	m_orderTable.Resize(2, nWeeks);;

	m_freeHR.Resize(3, nWeeks);;
	m_doingHR.Resize(3, nWeeks);;
	m_doneHR.Resize(3, nWeeks);;

	m_doingTable.Resize(11, nWeeks);
	m_doneTable.Resize(11, nWeeks);
	m_defferTable.Resize(11, nWeeks);
}

void CCompany::PrintDBTitle()
{
	int lastWeek = m_pGlobalEnv->SimulationWeeks;

	CString strDBoardTitle[1][18] = {
		{ _T("주"), _T("누계"), _T("발주"),_T(""), _T("투입"), _T("HR_H"), _T("HR_M"), _T("HR_L"),
		_T(""),_T("여유"), _T("HR_H"), _T("HR_M"), _T("HR_L"), _T(""),_T("총원"), _T("HR_H"), _T("HR_M"), _T("HR_L") }
	};
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 2, 1, (CString*)strDBoardTitle, 18, 1); //세로로 출력
	m_pXl->SetRangeBorder(WS_NUM_DASHBOARD, 2, 1, 4, lastWeek + 1, xlContinuous, xlThin, RGB(0, 0, 0));
	m_pXl->SetRangeBorder(WS_NUM_DASHBOARD, 7, 1, 9, lastWeek + 1, xlContinuous, xlThin, RGB(0, 0, 0));
	m_pXl->SetRangeBorder(WS_NUM_DASHBOARD, 12, 1, 14, lastWeek + 1, xlContinuous, xlThin, RGB(0, 0, 0));
	m_pXl->SetRangeBorder(WS_NUM_DASHBOARD, 17, 1, 19, lastWeek + 1, xlContinuous, xlThin, RGB(0, 0, 0));

	// 
	/*m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 2, 2, m_manageTable.pWeeksNum, 1, lastWeek);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 3, 2, m_manageTable.pSum, 1, lastWeek);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 4, 2, m_manageTable.pOrder, 1, lastWeek);*/

	int rows = m_orderTable.getRows();
	int cols = m_orderTable.getCols();

	int totalSize = rows * cols;  // Total number of elements
	int* tempBuf = new int[totalSize];  // Allocate memory for the total number of elements
	
	if ((ORDER_COUNT*lastWeek)!=totalSize)
	{
		MessageBox(NULL, _T("버퍼 사이즈를 확인하세요"), _T("Error"), MB_OK | MB_ICONERROR);
		return;
	}
	m_orderTable.copyToContinuousMemory(tempBuf, totalSize);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 3, 2, tempBuf, 4, lastWeek);

}