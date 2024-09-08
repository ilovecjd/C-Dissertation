
#include "stdafx.h"
#include "GlobalEnv.h"
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
	{
		m_pXl->ReleaseExcel();
		delete m_pXl;         // CXLEzAutomation 메모리 해제
	}

	if (m_pActType != NULL)
		delete m_pActType;    // PACT_TYPE 메모리 해제

	if (m_pActPattern != NULL)
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

	com_var.m_ID = Id;
	/////////////////////////////////////////////////////////////////////////
	// 전달 받은 환경 변수를 Company 로 복사
	*m_pGlobalEnv = *pGlobalEnv;	

	//m_pXl->ReadRangeToArray(WS_NUM_ACTIVITY_STRUCT, 3, 2, (int*)m_pActType, 5, 13);
	//m_pXl->ReadRangeToArray(WS_NUM_ACTIVITY_STRUCT, 15, 2, (int*)m_pActPattern, 6, 26);
	*m_pActType = *m_pGlobalEnv->ActType;
	*m_pActPattern = *m_pGlobalEnv->ActPattern;
	
	// !!!!!! song --> 프로그램 종료시 배열들의 크기 동적으로 바뀐적이 있는지는 체크하는 루틴을 꼭 넣자
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

	// VARIANT 배열 생성 하고 VT_EMPTY로 초기화
	VARIANT projectInfo[iHeight][iWidth];  
	
	for (int i = 0; i < iHeight; ++i) {
		for (int j = 0; j < iWidth; ++j) {
			VariantInit(&projectInfo[i][j]);
			projectInfo[i][j].vt = VT_EMPTY;
		}
	}

	// 첫 번째 행 설정	
	posX = 0; posY = 0;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_category;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_ID;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_duration;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_startAvail;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_endDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_orderDate;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = static_cast<int>(pProject->prj_var.m_profit);
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_experience;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_winProb;

	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_nCashFlows;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_cashFlows[0];
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_cashFlows[1];
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_cashFlows[2];

	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_firstPay;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_secondPay;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_finalPay;

	
	// 두 번째 행 설정
	posX = 0; posY = 1;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.numActivities;

	posX = 10;  // 빈 칸을 건너뛰기
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_firstPayMonth;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_secondPayMonth;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_finalPayMonth;
	
	posX = 14;  // 빈 칸을 건너뛰기
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_projectType;
	projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activityPattern;
	
	// 활동 데이터 설정
	for (int i = 0; i < pProject->prj_var.numActivities; ++i) {
		// 인덱스를 문자열로 변환하고 "Activity" 접두사 추가
		CString strAct;		
		strAct.Format(_T("Activity%02d"), i + 1);

		posX = 1; // 엑셀의 2행 2열부터 적는다.
		projectInfo[posY][posX].vt = VT_BSTR; projectInfo[posY][posX++].bstrVal = strAct.AllocSysString();
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activities[i].duration;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activities[i].startDate;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activities[i].endDate;

		posX = 6;  // 두 열 건너뛰기
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activities[i].highSkill;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activities[i].midSkill;
		projectInfo[posY][posX].vt = VT_I4; projectInfo[posY][posX++].intVal = pProject->prj_var.m_activities[i].lowSkill;

		posY++;
	}

	int printY = 4 + (pProject->prj_var.m_ID -1)*iHeight;
	m_pXl->WriteArrayToRange(sheet, printY, 1, (VARIANT*)projectInfo, iHeight, iWidth);
	m_pXl->SetRangeBorderAround(sheet, printY, 1, printY + iHeight-1, iWidth + 1 - 1, 1, 2, RGB(0, 0, 0));
}


BOOL CCompany::CreateProjects()
{
	int cnt = 0, sum = 0;
	int lastWeek = m_pGlobalEnv->SimulationWeeks;

	/////////////////////////////////////////////////////////////////////////
	// 프로젝트 발주(발생) 현황 생성
	for (int week = 0; week < lastWeek; week++)
	{
		cnt = PoissonRandom(m_pGlobalEnv->WeeklyProb);	// 이번주 발생하는 프로젝트 갯수		
		m_orderTable[ORDER_SUM][week] = sum;			// 누계
		m_orderTable[ORDER_ORD][week] = cnt;			// 발생 프로젝트갯수
		sum = sum + cnt;	// 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다.
	}
	com_var.m_totalProjectNum = sum;
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
	m_pXl->ReadRangeToArray(WS_NUM_DASHBOARD, 3, 2, m_orderTable[0], 1, lastWeek);
	m_pXl->ReadRangeToArray(WS_NUM_DASHBOARD, 4, 2, m_orderTable[1], 1, lastWeek);
	
	com_var.m_totalProjectNum = m_orderTable[ORDER_SUM][lastWeek-1] + m_orderTable[ORDER_ORD][lastWeek-1];
	
	/////////////////////////////////////////////////////////////////////////
	// project 시트에 헤더 출력
	
	/////////////////////////////////////////////////////////////////////////
	// 프로젝트 생성 후 내용은 로드
	// song ==> project 의 생성자와 소멸자, init 함수를 확인해 놓자.
	
	// song ==> NULL 체크 하자
	m_AllProjects = new CProject*[com_var.m_totalProjectNum];
	int* pProjectInfo;

	LONG lInfoSize = 7 * com_var.m_totalProjectNum * 16;
	pProjectInfo = new int[lInfoSize];
	m_pXl->ReadExRangeConvertInt(WS_NUM_PROJECT, 4, 1, pProjectInfo, com_var.m_totalProjectNum * 7, 16);

	for (int i = 0; i < com_var.m_totalProjectNum; i++)
	{	
		LONG lBaseAddress = 0;
		LONG lTemp = 0;
		
		lBaseAddress = 7* i * 16;

		CProject* pTempPrj;
		pTempPrj = new CProject;
		
		// 첫 번째 행 설정			
		pTempPrj->prj_var.m_category		= *(pProjectInfo+lBaseAddress++);
		pTempPrj->prj_var.m_ID				= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_duration		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_startAvail		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_endDate			= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_orderDate		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_profit			= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_experience		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_winProb			= *(pProjectInfo + lBaseAddress++);
		
		pTempPrj->prj_var.m_nCashFlows		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_cashFlows[0]	= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_cashFlows[1]	= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_cashFlows[2]	= *(pProjectInfo + lBaseAddress++);
		
		pTempPrj->prj_var.m_firstPay		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_secondPay		= *(pProjectInfo + lBaseAddress++);
		pTempPrj->prj_var.m_finalPay		= *(pProjectInfo + lBaseAddress++);

		// 두 번째 행 
		lTemp = lBaseAddress;

		pTempPrj->prj_var.numActivities		= *(pProjectInfo + lBaseAddress++);
		
		// 활동 데이터 설정
		for (int j = 0; j < pTempPrj->prj_var.numActivities; j++)
		{
			lBaseAddress += 1;// 빈 칸을 건너뛰기 (Activity01)
			pTempPrj->prj_var.m_activities[j].duration = *(pProjectInfo + lBaseAddress++);
			pTempPrj->prj_var.m_activities[j].startDate = *(pProjectInfo + lBaseAddress++);
			pTempPrj->prj_var.m_activities[j].endDate = *(pProjectInfo + lBaseAddress++);
			
			lBaseAddress += 1;  // 빈 칸을 건너뛰기
			pTempPrj->prj_var.m_activities[j].highSkill = *(pProjectInfo + lBaseAddress++);
			pTempPrj->prj_var.m_activities[j].midSkill = *(pProjectInfo + lBaseAddress++);
			pTempPrj->prj_var.m_activities[j].lowSkill = *(pProjectInfo + lBaseAddress++);

			if (j == 0)
			{
				lBaseAddress += 1;  // 빈 칸을 건너뛰기
				pTempPrj->prj_var.m_firstPayMonth = *(pProjectInfo + lBaseAddress++);
				pTempPrj->prj_var.m_secondPayMonth = *(pProjectInfo + lBaseAddress++);
				pTempPrj->prj_var.m_finalPayMonth = *(pProjectInfo + lBaseAddress++);

				lBaseAddress += 1;  // 빈 칸을 건너뛰기
				pTempPrj->prj_var.m_projectType = *(pProjectInfo + lBaseAddress++);
				pTempPrj->prj_var.m_activityPattern = *(pProjectInfo + lBaseAddress++);

				lBaseAddress += 1;  // 빈 칸을 건너뛰기
			}
			else 
			{
				lBaseAddress += 8;  // 빈 칸을 건너뛰기
			}
		}

		pTempPrj->prj_var.m_isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
		
		m_AllProjects[i] = pTempPrj;
		// 디버깅 때만 사용 ==> PrintProjectInfo(WS_NUM_DEBUG_INFO, pTempPrj);
		lBaseAddress = lTemp + 6 * 16;
	}
	
	delete[] pProjectInfo;
	pProjectInfo = NULL;

	return TRUE;
}

// 이번 기간에 결정할 일들. 프로젝트의 신규진행, 멈춤, 인원증감 결정
BOOL CCompany::Decision(int thisWeek ) {

	com_var.m_lastDecisionWeek = thisWeek;
	// 1. 지난주에 진행중인 프로젝트중 완료되지 않은 프로젝트들만 이번주로 이관
	if (FALSE == CheckLastWeek(thisWeek))
	{
		//파산		
		//return FALSE;
	}

	// 2. 진행 가능한 후보프로젝트들을  candidateTable에 모은다
	SelectCandidates(thisWeek);

	// 3. 신규 프로젝트 선택 및 진행프로젝트 업데이트
	// 이번주에 발주된 프로젝트중 시작할 것이 있으면 진행 프로젝트로 기록
	SelectNewProject(thisWeek);

	PrintDBData();
	return TRUE;
}

// 완료프로젝트 검사 및 진행프로젝트 업데이트
// 1. 지난 기간의 정보를 이번기간에 복사하고
// 2. 지난 기간에 진행중인 프로젝트중 완료된 것이 있는가?
// 3. 완료된 프로젝트들만 이번기간에서 삭제
BOOL CCompany::CheckLastWeek(int thisWeek )
{	
	if (0 == thisWeek) // 첫주는 체크할 지난주가 없음
		return TRUE;

	int nLastProjects = m_doingTable[ORDER_SUM][thisWeek - 1];//지난주에 진행 중이던 프로젝트의 갯수
	
	for (int i = 0; i < nLastProjects; i++)
	{
		int prjId = m_doingTable[i + 1][thisWeek - 1];
		if (prjId == 0)
			return TRUE;

		CProject* project = m_AllProjects[prjId-1];

		// 1. payment 를 계산한다. 선금은 시작시 받기로 한다. 조건완료후 1주 후 수금			
		// 2. 지출을 계산한다.
		//' 3. 진행중인 프로젝트를 이관해서 기록한다.
		int sum = m_doingTable[ORDER_SUM][thisWeek];
		if (thisWeek < (project->prj_var.m_isStart + project->prj_var.m_duration - 1)) // ' 아직 안끝났으면
		{			
			m_doingTable[sum + 1][thisWeek] = project->prj_var.m_ID;// 테이블 크기는 자동으로 변경된다.
			m_doingTable[ORDER_SUM][thisWeek] = sum + 1;
		}
	}

	// 자금 현황을 체크하자.
	// 나중에 후회 해도 일단은 편하게 코딩.
	int Cash = m_pGlobalEnv->Cash_Init;

	for (int i = 0; i < thisWeek; i++)
	{
		Cash += (m_incomeTable[0][i] - m_expensesTable[0][i]);
	}
	//if (Cash<0)// 파산
	//{
	//	return FALSE;
	//}

	if (1003 < thisWeek)
	{
		/// 인원 증감을 결정하자.
		int temp = m_expensesTable[0][thisWeek] * m_pGlobalEnv->recruit;
		if (temp < Cash)
		{
			int i = rand() % 3;
			AddHR(i, thisWeek + m_pGlobalEnv->Hr_LeadTime);// 인원 충원 리드 타임
		}

		temp = m_expensesTable[0][thisWeek] * m_pGlobalEnv->layoff;
		if (temp > Cash)
		{
			int i = rand() % 3;
			RemoveHR(i, thisWeek + m_pGlobalEnv->Hr_LeadTime);// 인원 감원 리드 타임
		}
	}
	
	return TRUE;
}

void CCompany::AddHR(int grade ,int addWeek)
{
	// 충원 / 감원 인원 추가
	// 나머지 기간 업데이트
	// 나머지 기간의 비용 업데이트
	m_totalHR[grade][addWeek] = m_totalHR[grade][addWeek] + 1;

	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
	double rate = m_pGlobalEnv->ExpenseRate;
	int expenses = (m_totalHR[0][addWeek] * 50 * rate) + (m_totalHR[1][addWeek] * 39 * rate) + (m_totalHR[2][addWeek] * 25 * rate);

	for (int i = addWeek; i < m_pGlobalEnv->SimulationWeeks + ADD_HR_SIZE; i++)
	{
		m_totalHR[HR_HIG][i] = m_totalHR[HR_HIG][addWeek];
		m_totalHR[HR_MID][i] = m_totalHR[HR_MID][addWeek];
		m_totalHR[HR_LOW][i] = m_totalHR[HR_LOW][addWeek];
		m_expensesTable[0][i] = expenses;
	}
}
//
void CCompany::RemoveHR(int grade, int removeWeek)
{
	// 감원 인원 
	// 나머지 기간 업데이트
	// 나머지 기간의 비용 업데이트
	m_totalHR[grade][removeWeek] = m_totalHR[grade][removeWeek] + 1;

	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
	double rate = m_pGlobalEnv->ExpenseRate;
	int expenses = (m_totalHR[0][removeWeek] * 50 * rate) + (m_totalHR[1][removeWeek] * 39 * rate) + (m_totalHR[2][removeWeek] * 25 * rate);

	for (int i = removeWeek; i < m_pGlobalEnv->SimulationWeeks + ADD_HR_SIZE; i++)
	{
		m_totalHR[HR_HIG][i] = m_totalHR[HR_HIG][removeWeek];
		m_totalHR[HR_MID][i] = m_totalHR[HR_MID][removeWeek];
		m_totalHR[HR_LOW][i] = m_totalHR[HR_LOW][removeWeek];
		m_expensesTable[0][i] = expenses;
	}
}


void CCompany::SelectCandidates(int thisWeek)
{
	int lastID = m_orderTable[ORDER_SUM][thisWeek] ;	// 지난달까지 누계
	int endID = m_orderTable[ORDER_ORD][thisWeek] + lastID;  // 지난달까지 누계 + 이번주 발생 갯수 - 1
	for(int i=0; i< MAX_CANDIDATES; i++)
		m_candidateTable[i] = 0;

	int j = 0; 
	for (int i = lastID; i < endID; i++)
	{
		CProject* project = m_AllProjects[i];

		if (IsEnoughHR(thisWeek, project)) // 인원 체크
		{
			m_candidateTable[j++] = project->prj_var.m_ID;
		}
	}
}

BOOL CCompany::IsEnoughHR(int thisWeek, CProject* project)
{
	// 원본 인력 테이블을 복사해서 프로젝트 인력을 추가 할 수 있는지 확인한다.
	int** doingHR = Newallocate2DArray(3, m_pGlobalEnv->SimulationWeeks+ ADD_HR_SIZE);
	Newcopy2DArray(m_doingHR, doingHR, 3, m_pGlobalEnv->SimulationWeeks+ ADD_HR_SIZE);
		
	// 2중 루프 activity->기간-> 등급업데이트 순서로 activity들을 순서대로 가져온다.
	int numAct = project->prj_var.numActivities;
	for (int i = 0 ; i < numAct ;i++)
	{
		PACTIVITY pActivity = &(project->prj_var.m_activities[i]);
		for (int j = 0; j < pActivity->duration; j++)
		{
			doingHR[HR_HIG][j + pActivity->startDate] += pActivity->highSkill;
			doingHR[HR_MID][j + pActivity->startDate] += pActivity->midSkill;
			doingHR[HR_LOW][j + pActivity->startDate] += pActivity->lowSkill;
		}		
	}

	for (int i = thisWeek; i < m_pGlobalEnv->SimulationWeeks; i++) 
	{
		if (m_totalHR[HR_HIG][i] < doingHR[HR_HIG][i])
		{
			Newdeallocate2DArray(doingHR, 3);
			return FALSE;
		}
			

		if (m_totalHR[HR_MID][i] < doingHR[HR_MID][i])
		{
			Newdeallocate2DArray(doingHR, 3);
			return FALSE;
		}

		if (m_totalHR[HR_LOW][i] < doingHR[HR_LOW][i])
		{
			Newdeallocate2DArray(doingHR, 3);
			return FALSE;
		}
	}

	Newdeallocate2DArray(doingHR, 3);
	return TRUE;
}

// 후보군들을 선택 정책에 따라서 순서를 변경한다.

// 2차원 배열을 오름차순으로 정렬하는 함수
void sortArrayAscending(int* indexArray, int* valueArray, int size) {
	// 두 배열을 정렬하기 위해 값과 인덱스를 페어로 묶어야 합니다.
	for (int i = 0; i < size - 1; i++) {
		for (int j = i + 1; j < size; j++) {
			if (valueArray[i] > valueArray[j]) {
				// 값(value)을 기준으로 정렬하고, 인덱스도 함께 변경합니다.
				std::swap(valueArray[i], valueArray[j]);
				std::swap(indexArray[i], indexArray[j]);
			}
		}
	}
}

// 2차원 배열을 내림차순으로 정렬하는 함수
void sortArrayDescending(int* indexArray, int* valueArray, int size) {
	for (int i = 0; i < size - 1; i++) {
		for (int j = i + 1; j < size; j++) {
			if (valueArray[i] < valueArray[j]) {
				// 값(value)을 기준으로 내림차순으로 정렬하고, 인덱스도 함께 변경합니다.
				std::swap(valueArray[i], valueArray[j]);
				std::swap(indexArray[i], indexArray[j]);
			}
		}
	}
}
void CCompany::SelectNewProject(int thisWeek)
{	
	
	int valueArray[MAX_CANDIDATES] = {0, };  // 값 배열
	int j = 0;

	while (m_candidateTable[j] != 0) {
		int id = m_candidateTable[j];
		CProject* project = m_AllProjects[id - 1];
		valueArray[j] = project->prj_var.m_profit;
		j = j + 1;
	}
	
	//switch (m_pGlobalEnv->selectOrder)
	//{
	//case 1: // 발생 순서대로
	//	break;
	//case 2:
	//	sortArrayAscending(m_candidateTable, valueArray, j);	// 금액 내림차순 정렬	
	//	break;
	//case 3:
	//	sortArrayDescending(m_candidateTable, valueArray, j); // 금액 오름차순 정렬	
	//	break;
	//default : 
	//	break;
	//}
	
	


	int i = 0;
	while(m_candidateTable[i] != 0) {

		if (i > MAX_CANDIDATES) break;

		int id = m_candidateTable[i++];

		CProject* project = m_AllProjects[id-1];

		if (project->prj_var.m_startAvail < m_pGlobalEnv->SimulationWeeks)
		{
			if (IsEnoughHR(thisWeek,project))
			{	
				AddProjectEntry(project, thisWeek);	

				///int lows = m_debugInfo.getRows();
				
				int tempTotal = project->prj_var.m_firstPay + project->prj_var.m_secondPay + project->prj_var.m_finalPay;

				int cols = m_debugInfo.getCols();
				m_debugInfo.Resize(2, cols+1);

				m_debugInfo[0][cols] = project->prj_var.m_ID;
				m_debugInfo[1][cols] = tempTotal;

			}
		}		
	} 
}

// 모든 체크가 끝나고 호출된다. 
// 단지 변수들만 셑팅하자.
void CCompany::AddProjectEntry(CProject* project,  int addWeek)
{	
	project->prj_var.m_isStart = project->prj_var.m_startAvail;

	// HR 정보 업데이트
	// 2중 루프 activity->기간-> 등급업데이트 순서로 activity들을 순서대로 가져온다.
	int numAct = project->prj_var.numActivities;
	for (int i = 0; i < numAct; i++)
	{
		PACTIVITY pActivity = &(project->prj_var.m_activities[i]);
		for (int j = 0; j < pActivity->duration; j++)
		{
			int col = j + pActivity->startDate;
			m_doingHR[HR_HIG][col] += pActivity->highSkill;
			m_doingHR[HR_MID][col] += pActivity->midSkill;
			m_doingHR[HR_LOW][col] += pActivity->lowSkill;

			m_freeHR[HR_HIG][col] = m_totalHR[HR_HIG][col] - m_doingHR[HR_HIG][col];
			m_freeHR[HR_MID][col] = m_totalHR[HR_MID][col] - m_doingHR[HR_MID][col];
			m_freeHR[HR_LOW][col] = m_totalHR[HR_LOW][col] - m_doingHR[HR_LOW][col];
		}
	}

	// 현황판 업데이트
	int sum = m_doingTable[0][addWeek];
	m_doingTable[sum + 1][addWeek] = project->prj_var.m_ID;
	m_doingTable[0][addWeek] = sum + 1;

	// 수입 테이블 업데이트. 지출은 인원 관리쪽에서 한다.	
	int incomeDate;

	if (project->prj_var.m_isStart <addWeek)
	{
		MessageBox(NULL, _T("m_isStart miss"), _T("Error"), MB_OK | MB_ICONERROR);
	}
	incomeDate = project->prj_var.m_isStart + project->prj_var.m_firstPayMonth;	// 선금 지급일
	m_incomeTable[0][incomeDate] += project->prj_var.m_firstPay;
	
	incomeDate = project->prj_var.m_isStart + project->prj_var.m_secondPayMonth;	// 2차 지급일
	m_incomeTable[0][incomeDate] += project->prj_var.m_secondPay;

	incomeDate = project->prj_var.m_isStart + project->prj_var.m_finalPayMonth;	// 3차 지급일
	m_incomeTable[0][incomeDate] += project->prj_var.m_finalPay;
}


// dash boar 용 배열들의 크기 조절	
void CCompany::AllTableInit(int nWeeks)
{
	m_orderTable = Newallocate2DArray(2, nWeeks);

	m_doingHR = Newallocate2DArray(3, nWeeks + ADD_HR_SIZE);
	m_freeHR = Newallocate2DArray(3, nWeeks + ADD_HR_SIZE);
	m_totalHR = Newallocate2DArray(3, nWeeks + ADD_HR_SIZE);

	m_doingTable.Resize(11, nWeeks);
	m_doneTable.Resize(11, nWeeks);
	m_defferTable.Resize(11, nWeeks);

	m_incomeTable = Newallocate2DArray(1, nWeeks);
	m_expensesTable = Newallocate2DArray(1, nWeeks);


	// 이건 충원이나 감원쪽에서 필요시 다시 수정하게 된다.	
	m_totalHR[HR_HIG][0] = m_freeHR[HR_HIG][0] = m_pGlobalEnv->Hr_Init_H;
	m_totalHR[HR_MID][0] = m_freeHR[HR_MID][0] = m_pGlobalEnv->Hr_Init_M;
	m_totalHR[HR_LOW][0] = m_freeHR[HR_LOW][0] = m_pGlobalEnv->Hr_Init_L;

	// 소요 비용 계산. 수정시 다음도 수정 필요 CProject::CalculateLaborCost(const std::string& grade)
	double rate = m_pGlobalEnv->ExpenseRate;
	int expenses = (m_pGlobalEnv->Hr_Init_H * 50* rate) + (m_pGlobalEnv->Hr_Init_M * 39* rate) + (m_pGlobalEnv->Hr_Init_L * 25 * rate);

	for (int i = 0; i < nWeeks + ADD_HR_SIZE; i++)
	{
		m_totalHR[HR_HIG][i] = m_pGlobalEnv->Hr_Init_H;
		m_totalHR[HR_MID][i] = m_pGlobalEnv->Hr_Init_M;
		m_totalHR[HR_LOW][i] = m_pGlobalEnv->Hr_Init_L;
		m_expensesTable[0][i] = expenses;
	}
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

	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 3, 2, m_orderTable[0], 1, m_pGlobalEnv->SimulationWeeks);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 4, 2, m_orderTable[1], 1, m_pGlobalEnv->SimulationWeeks);

	int* pWeeks = new int[m_pGlobalEnv->SimulationWeeks];
	for (int i = 0; i < m_pGlobalEnv->SimulationWeeks; i++)
	{
		pWeeks[i] = i + 1;
	}
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 2, 2, pWeeks, 1, m_pGlobalEnv->SimulationWeeks);

}


void CCompany::PrintDBData()
{
	int rows = m_debugInfo.getRows();
	int cols = m_debugInfo.getCols();

	int totalSize = rows * cols;  // Total number of elements
	int* tempBuf = new int[totalSize];  // Allocate memory for the total number of elements

	// cash flow
	m_debugInfo.copyToContinuousMemory(tempBuf, totalSize);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 7, 2, tempBuf, rows, cols);

	delete[] tempBuf;

	// 다같은 사이즈 이니 한번만 계산해서 사용하자

	cols = m_pGlobalEnv->SimulationWeeks + ADD_HR_SIZE;
	
	// HR 정보 출력	
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 7 + 3, 2, m_doingHR[0], 1, cols);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 7 + 4, 2, m_doingHR[1],1, cols);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 7 + 5, 2, m_doingHR[2], 1, cols);

	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 12 + 3, 2, m_freeHR[0], 1, cols);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 12 + 4, 2, m_freeHR[1], 1, cols);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 12 + 5, 2, m_freeHR[2], 1, cols);

	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 17 + 3, 2, m_totalHR[0], 1, cols);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 17 + 4, 2, m_totalHR[1], 1, cols);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 17 + 5, 2, m_totalHR[2], 1, cols);
	
	int printRow = 22 + 3;
	// 진행 현황 출력
	rows = m_doingTable.getRows();
	cols = m_doingTable.getCols();
	totalSize = rows * cols;  // Total number of elements
	tempBuf = new int[totalSize];  // Allocate memory for the total number of elements

	m_doingTable.copyToContinuousMemory(tempBuf, totalSize);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, printRow, 2, tempBuf, rows, cols);
	printRow += rows + 1;
	delete[] tempBuf;

	////////////////////////////////////////////////	
	rows = m_doneTable.getRows();
	cols = m_doneTable.getCols();
	totalSize = rows * cols;  // Total number of elements
	tempBuf = new int[totalSize];  // Allocate memory for the total number of elements

	m_doneTable.copyToContinuousMemory(tempBuf, totalSize);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, printRow, 2, tempBuf, rows, cols);
	printRow += rows + 1;
	delete[] tempBuf;

	////////////////////////////////////////////////
	rows = m_defferTable.getRows();
	cols = m_defferTable.getCols();
	totalSize = rows * cols;  // Total number of elements
	tempBuf = new int[totalSize];  // Allocate memory for the total number of elements

	m_defferTable.copyToContinuousMemory(tempBuf, totalSize);
	m_pXl->WriteArrayToRange(WS_NUM_DASHBOARD, printRow, 2, tempBuf, rows, cols);

	delete[] tempBuf;
}

int CCompany::CalculateFinalResult() 
{
	int result = m_pGlobalEnv->Cash_Init;

	for (int i = 0; i < com_var.m_lastDecisionWeek; i++)
	{
		result += (m_incomeTable[0][i]- m_expensesTable[0][i]);
	}
	
/*	무슨코드였는지 모르겠다. 다음에 정리하자.
	int tempTotalIncome = m_pGlobalEnv->Cash_Init;
	int tempOutcome = m_expensesTable[0][10];
	int tempTetoalOutcome = (tempOutcome*144);
	
	int cols = m_debugInfo.getCols();
	for (int i = 0; i < cols; i++)
	{		
		tempTotalIncome += m_debugInfo[1][i];
	}
	int tempResult = tempTotalIncome- tempTetoalOutcome;
	*/
	//return result;
	//return tempResult; // 기대수익?? 포함 (수주 수익 포함)	
	return result;
}



/*
1. 시그니처 SONG1 / 2.전체크기
3. 환경변수
4. 엑티비티타입
5. 엑티비티패턴
6. 현황판
프로젝트갯수
프로젝트
*/


#define SIGNITURE		{'A','H','N','1'} //pack 사용하지 않게 4바이트 정렬
#define TYPE_UNKNOWN		0
#define TYPE_ANH			1
#define TYPE_ENVIRANMENT	2
#define TYPE_ACTIVITY		3
#define TYPE_PATTERN		4
#define TYPE_ORDER			5
#define TYPE_DASHBD			6
#define TYPE_PROJECT		7

struct SAVE_SIG
{
	char signitre[4] = SIGNITURE;
	int totalLen;
};

struct SAVE_TL
{
	int type;
	int length;
};


bool OpenFile(const CString& filename, const TCHAR* mode, FILE** fp) {
	errno_t err = _wfopen_s(fp, filename, mode);
	if (err != 0 || *fp == nullptr) {
		perror("Failed to open file");
		return false;
	}
	return true;
}

void CloseFile(FILE** fp) {
	if (*fp != nullptr) {
		fclose(*fp);
		*fp = nullptr;
	}
}


void WriteDataWithHeader(FILE* fp, int type, const void* data, size_t dataSize) {
	SAVE_TL tl = { type, static_cast<int>(dataSize) };
	fwrite(&tl, sizeof(tl), 1, fp);  // 먼저 데이터 타입 및 길이 정보를 쓴다
	fwrite(data, dataSize, 1, fp);   // 실제 데이터 쓰기
}


void CCompany::SaveProjectToAhn(const CString& filename) {
	FILE* fp = nullptr;
	if (!OpenFile(filename, _T("wb"), &fp)) return;

	SAVE_SIG sig;
	fwrite(&sig, sizeof(sig), 1, fp);  // 파일 시작 부분에 시그니처 쓰기
		
	WriteDataWithHeader(fp, TYPE_ENVIRANMENT, m_pGlobalEnv, sizeof(GLOBAL_ENV));
	WriteDataWithHeader(fp, TYPE_ACTIVITY, m_pActType, sizeof(ALL_ACT_TYPE));
	WriteDataWithHeader(fp, TYPE_PATTERN, m_pActPattern, sizeof(ALL_ACTIVITY_PATTERN));
	WriteDataWithHeader(fp, TYPE_ORDER, m_orderTable, 2 * m_pGlobalEnv->SimulationWeeks);
	
	sig.totalLen = 0;
	fclose(fp);
	fp = nullptr;
	// 파일 닫기
	CloseFile(&fp);
}


//
//
//void CCompany::LoadProjectFromAhn(const CString& filename)
//{	
//	FILE* fp = nullptr;
//	if (!OpenFile(filename, _T("rb"), &fp)) return;
//
//	SAVE_SIG sig;
//	ReadData(fp, &sig, sizeof(sig));
//
//	long fileSize = ftell(fp); // 파일 크기를 얻음
//	rewind(fp); // 파일 포인터를 다시 파일 시작으로 이동
//
//	GLOBAL_ENV env;
//	if (LoadData(fp, TYPE_ENVIRONMENT, &m_pGlobalEnv, sizeof(env)));
//
//
//	SAVE_TL tl;
//
//	errno_t err = _wfopen_s(&fp, filename, _T("rb"));  // _wfopen_s는 안전한 함수입니다
//
//	if (err != 0 || fp == nullptr) {  // 오류가 발생했거나 파일 포인터가 null인 경우
//		perror("Failed to open file for writing");
//		return;
//	}
//
//	//// 파일의 크기를 알아내기 위해 파일 포인터를 파일 끝으로 이동
//	fseek(fp, 0, SEEK_END);
//
//	if (fseek(fp, 0, SEEK_END) != 0) {
//		perror("Failed to seek to end of file");
//		fclose(fp);
//		return ;
//	}
//
//	
//	// 파일 크기만큼의 메모리를 동적 할당
//	unsigned char* buffer = (unsigned char*)malloc(fileSize);
//	if (buffer == nullptr) {
//		perror("Failed to allocate memory");
//		fclose(fp);
//		return ;
//	}
//
//	// 파일 내용을 버퍼에 읽어옴
//	size_t bytesRead = fread(buffer, sizeof(unsigned char), fileSize, fp);
//	if (bytesRead != fileSize) {
//		perror("Failed to read the entire file");
//		free(buffer);
//		fclose(fp);
//		return ;
//	}
//
//
//	SAVE_SIG sig;
//	ULONG bufPos = 0; // 읽어야할 버퍼의 위치, 계속 증가 시킴
//	ULONG ulSize = 0;
//	
//	ulSize = sizeof(SAVE_SIG);
//
//	if (memcmp(sig.signitre, buffer, 4) == 0) {
//		perror("The first 4 bytes of signitre and buffer are the same.");
//	}
//
//	// 1. 시그니처
//	memcpy(&sig, buffer, ulSize);
//	bufPos += ulSize;
//	
//	// 2. 환경변수
//	ulSize = sizeof(SAVE_TL);
//	memcpy(&tl, buffer+bufPos, ulSize);
//	bufPos += ulSize;
//
//	ulSize = tl.length;
//	if (TYEP_ENVIRANMENT == tl.type )
//	{ 
//		if (sizeof(GLOBAL_ENV) == ulSize)
//		{
//			memcpy(m_pGlobalEnv, buffer + bufPos, ulSize);
//			bufPos += ulSize;
//		}
//	}
//
//	// 3. 
//	ulSize = sizeof(SAVE_TL);
//	memcpy(&tl, buffer + bufPos, ulSize);
//	bufPos += ulSize;
//
//	ulSize = tl.length;
//	if (TYPE_ACTIVITY == tl.type)
//	{
//		if (sizeof(ALL_ACT_TYPE) == ulSize)
//		{
//			memcpy(m_pActType, buffer + bufPos, ulSize);
//			bufPos += ulSize;
//		}
//	}
//
//	// 4.
//
//	ulSize = sizeof(SAVE_TL);
//	memcpy(&tl, buffer + bufPos, ulSize);
//	bufPos += ulSize;
//
//	ulSize = tl.length;
//	if (TYPE_PATTERN == tl.type)
//	{
//		if (sizeof(ALL_ACTIVITY_PATTERN) == ulSize)
//		{
//			memcpy(m_pActPattern, buffer + bufPos, ulSize);
//			bufPos += ulSize;
//		}
//	}
//
//	// 5. order table
//	ulSize = sizeof(SAVE_TL);
//	memcpy(&tl, buffer + bufPos, ulSize);
//	bufPos += ulSize;
//
//	ulSize = tl.length;
//	if (TYPE_ORDER == tl.type)
//	{
//		if (m_pGlobalEnv->SimulationWeeks*2 == ulSize)
//		{
//			if(NULL == m_orderTable )
//				m_orderTable = malloc(ulSize);
//			memcpy(m_orderTable, buffer + bufPos, ulSize);
//			bufPos += ulSize;
//		}
//	}
//
//
//	// 6. projects
//	if (TYPE_PATTERN == tl.type)
//	//{
//	//	if (sizeof(ALL_ACTIVITY_PATTERN) == tl.length)
//	//	{
//	//		m_orderTable.copyFromContinuousMemory((int *)buffer, 2, ulSize / 4 / 2);
//	//		bufPos += ulSize;
//	//	}
//	//}
//
//	tl.type = TYPE_DASHBD;
//	//tl.length = totalSize;
//	//fwrite(&tl, sizeof(SAVE_TL), 1, fp);
//
//	//fwrite(tempBuf, totalSize, 1, fp);
//
//	//#define TYPE_PROJECT		6
//
//	///***** 전체 크기는 마지막에 계산해서 넣어주자.
//	//sig.totalLen = 0;
//	
//
//	// 메모리 해제 및 파일 닫기
//	free(buffer);
//	fclose(fp);//	fp = nullptr;
//}
