﻿#include "stdafx.h"
#include "GlobalEnv.h"
#include "Creator.h"
#include <cctype>   // toupper 함수를 사용하기 위해 필요

CCreator::CCreator()
{
	
}

CCreator::~CCreator()
{	
}

// 문제를 생성한다. 글로벌 환경에서 발생 할 수 있는 프로젝트들을 발생 시키고 파일로 저장한다.
//BOOL CCreator::Init(int type, int ID, int ODate, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern)
BOOL CCreator::Init(GLOBAL_ENV* pGlobalEnv, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern)
{	
	*(&m_GlobalEnv) = *pGlobalEnv;
	*(&m_ActType) = *pActType;
	*(&m_ActPattern) = *pActPattern;

	CreateOrderTable();//m_totalProjectNum 생성 (내부프로젝트 3개 만큼 크게)
	m_pProjects = new PROJECT[m_totalProjectNum];
	CreateProjects();



	//prj_var.m_category		= type;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	//prj_var.m_ID			= ID;		// 프로젝트의 번호	
	//prj_var.m_orderDate		= ODate;	// 발주일
	//prj_var.m_startAvail	= ODate + (rand() % 4);  // // 시작 가능일 ( 0에서 3 사이의 정수 난수 생성)
	//prj_var.m_isStart		= 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
	//prj_var.m_experience	= ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
	//prj_var.m_winProb		= 100;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 현재는 100%
	//prj_var.m_nCashFlows	= MAX_N_CF;	// 비용 지급 횟수(규모에 따라 변경 가능)

	//CreateActivities();					//m_activities[MAX_ACT] 계산
	//prj_var.m_profit = CalculateHRAndProfit(); // 총 수익을 계산한다.
	//CalculatePaymentSchedule();			//m_cashFlows[MAX_N_CF] 계산
	return TRUE;
}

/////////////////////////////////////////////////////////////////////////
// 프로젝트 발주(발생) 현황 생성, 프로젝트는 최대 크기 만큼 설정한다.
// 시뮬레이션보다 길게 작성한다.
int CCreator::CreateOrderTable()
{
	int cnt = 0, sum = 0;

	m_orderTable.Resize(2, m_GlobalEnv.maxWeek);

	for (int week = 0; week < m_GlobalEnv.maxWeek; week++)
	{
		cnt = PoissonRandom(m_GlobalEnv.WeeklyProb);	// 이번주 발생하는 프로젝트 갯수		
		m_orderTable[ORDER_SUM][week] = sum;			// 누계
		m_orderTable[ORDER_ORD][week] = cnt;			// 발생 프로젝트갯수
		sum = sum + cnt;	// 이번주 까지 발생한 프로젝트 갯수. 다음주에 기록된다.
	}
	m_OutProjectNum = sum;
	m_totalProjectNum = sum + 3;// 생성될 내부프로젝트 최대 갯수만큼 더한다.
	return 0;
}

int CCreator::CreateProjects()
{
	int projectId = 0;
	int startNum = 0;
	int endNum = 0;
	int preTotal = 0;

	// 외부 프로젝트 생성
	for (int week = 0; week < m_GlobalEnv.maxWeek; week++)
	{
		preTotal = m_orderTable[ORDER_SUM][week];			// 지난주까지의 발주 프로젝트 누계
		startNum = preTotal + 1;						// 신규프로젝트이 시작번호 = 누계 +1
		endNum = preTotal + m_orderTable[ORDER_ORD][week];	// 마지막 프로젝트의 시작번호 = 지난주 누계 + 이번주 발생건수

		if ((startNum != 0) && (startNum <= endNum))
		{
			for (projectId = startNum; projectId <= endNum; projectId++)
			{
				PROJECT* pProject;				
				pProject =  &m_pProjects[projectId-1];
				memset(pProject, 0, sizeof(struct PROJECT));

				pProject->category = 0;		// 프로젝트 분류 (0: 외부 / 1: 내부)
				pProject->ID = projectId;		// 프로젝트의 번호	
				pProject->orderDate = week;	// 발주일
				pProject->startAvail = week  + (rand() % 4);  // // 시작 가능일 ( 0에서 3 사이의 정수 난수 생성)
				pProject->isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
				pProject->experience = ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
				pProject->winProb = 100;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 현재는 100%
				pProject->nCashFlows = MAX_N_CF;	// 비용 지급 횟수(규모에 따라 변경 가능)

				CreateActivities(pProject);					//m_activities[MAX_ACT] 계산
				pProject->profit = CalculateHRAndProfit(pProject); // 총 수익을 계산한다.
				CalculatePaymentSchedule(pProject);			//m_cashFlows[MAX_N_CF] 계산				
			}
		}
	}

	// 내부 프로젝트 생성
	for (int i = 0; i <3; i++)
	{	
		PROJECT* pProject;
		int duration = 40+i*4;  // 10개월, 11개월, 12개월
		int startDate = i * 4 * 12;

		projectId = m_OutProjectNum + 1 + i;
		pProject = &m_pProjects[projectId-1];
		memset(pProject, 0, sizeof(struct PROJECT));

		pProject->category = 1;		// 프로젝트 분류 (0: 외부 / 1: 내부)
		pProject->ID = projectId;		// 프로젝트의 번호	
		pProject->orderDate = startDate;	// 발주일
		pProject->startAvail = startDate; // 바로시작 가능
		pProject->isStart = 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
		pProject->experience = ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
		pProject->winProb = 30 + i*10;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 
		pProject->nCashFlows = 0;			// 비용 지급 횟수(규모에 따라 변경 가능)
		
		pProject->endDate = startDate + duration - 1;		// 프로젝트 종료일
		pProject->duration = duration;		// 프로젝트의 총 기간

		//CalculatePaymentSchedule(pProject);			//m_cashFlows[MAX_N_CF] 계산	
		pProject->profit = m_GlobalEnv.Cash_Init /12/4;	// 주당 기대수익으로 변경
															
		//1CreateActivities(pProject);			//m_activities[MAX_ACT] 계산 // 활동
		pProject->numActivities = 1;          // 총 활동 수
		pProject->activities[0].activityType = 1;// 활동에 관한 정보를 기록하는 배열	
		pProject->activities[0].duration = duration;      // 활동 기간
		pProject->activities[0].startDate = startDate;     // 시작 날짜
		pProject->activities[0].endDate = startDate + duration - 1;       // 종료 날짜
		pProject->activities[0].highSkill = m_GlobalEnv.Hr_Init_H / 2;     // 높은 기술 수준 인력 수
		pProject->activities[0].midSkill = m_GlobalEnv.Hr_Init_H / 2;      // 중간 기술 수준 인력 수
		pProject->activities[0].lowSkill = m_GlobalEnv.Hr_Init_H / 2;      // 낮은 기술 수준 인력 수

	}

	return 0;

}


BOOL CCreator::CreateActivities(PROJECT* pProject) {
	//song 사용하지 않는 멤버 변수와 지역 변수들 삭제 하자	
	int i;
	int probability;
	int Lb = 0;
	int UB = 0;
	int maxLoop;
	int totalDuration;
	int tempDuration;

	probability = rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
	maxLoop = MAX_PRJ_TYPE;

	////////////////////////////////////////////
	// m_pActType->asIntArray[][] : activity_struct 시트의 cells(3,2) ~ cells(7,14) 의 값이 들어 있음.
	
	////////////////////////////////////////////
	// 프로젝트 타입관련 정보	
	for (i = 0; i < maxLoop; ++i) { // 프로젝트 타입을 결정한다
		UB += m_ActType.asIntArray[i][0];	// 엑셀 2열의 "발생 확률"

		if (Lb <= probability && probability < UB) {
			pProject->projectType = i;
			break;
		}
		Lb = UB;
	}
	
	Lb = m_ActType.asIntArray[pProject->projectType][2];	// 엑셀 4열의 "최소기간"
	UB = m_ActType.asIntArray[pProject->projectType][3];	// 엑셀 5열의 "최대기간"

	totalDuration = RandomBetween(Lb, UB);
	pProject->duration = totalDuration;
	pProject->endDate = pProject->startAvail + totalDuration - 1;// song??

	Lb = 0;
	UB = 0;
	maxLoop = m_ActType.asIntArray[pProject->projectType][4];//패턴수

	// 패턴 타입 결정
	for (i = 0; i < maxLoop; ++i) {
		UB += m_ActType.asIntArray[pProject->projectType][6 + ((i) * 2)];//1번패턴 확률부터

		if (Lb <= probability && probability < UB) {
			pProject->activityPattern = m_ActType.asIntArray[pProject->projectType][5 + ((i) * 2)];//1번패턴 패턴번호부터
			break;
		}
		Lb = UB;
	}

	//////////////////////////////////////////////////////////////////
	//프로젝트 패턴 관련 정보
	Lb = 0;
	UB = 0;
	maxLoop = m_ActPattern.asIntArray[pProject->activityPattern-1][0];//활동수 !!! -1 에 주의
	pProject->numActivities = maxLoop;

	// 활동 생성
	for (i = 0; i < maxLoop; ++i) {
		Lb += m_ActPattern.asIntArray[pProject->activityPattern-1][1 + ((i) * 5)];// !!! -1 에 주의
		UB += m_ActPattern.asIntArray[pProject->activityPattern-1][2 + ((i) * 5)];// !!! -1 에 주의
		probability = RandomBetween(Lb, UB);
		tempDuration = totalDuration * probability / 100;

		if (tempDuration == 0) {
			tempDuration = 1;
		}

		if (i == 0) {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->startAvail;
			pProject->activities[i].endDate = pProject->startAvail - 1 + tempDuration;
		}
		else if (i == 1) {
			pProject->activities[i].duration = totalDuration - pProject->activities[0].duration;
			pProject->activities[i].startDate = pProject->activities[0].endDate + 1;
			pProject->activities[i].endDate = pProject->startAvail - 1 + totalDuration;
		}
		else if (i == 2) {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->startAvail - 1 + totalDuration - tempDuration + 1;
			pProject->activities[i].endDate = pProject->startAvail - 1 + totalDuration;
		}
		else {
			pProject->activities[i].duration = tempDuration;
			pProject->activities[i].startDate = pProject->activities[2].startDate - tempDuration;
			pProject->activities[i].endDate = pProject->activities[2].startDate - 1;
		}
	}
	return TRUE;
}

// 활동별 투입 인력 생성 및 프로젝트 전체 기대 수익 계산 함수
int CCreator::CalculateHRAndProfit(PROJECT* pProject) {
	int high = 0, mid = 0, low = 0;

	for (int i = 0; i < pProject->numActivities; ++i) {
		int j = rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
		if (0 < j && j <= RND_HR_H) {
			pProject->activities[i].highSkill	= 1;
			pProject->activities[i].midSkill	= 0;
			pProject->activities[i].lowSkill	= 0;
		}
		else if (RND_HR_H < j && j <= RND_HR_M) {
			pProject->activities[i].highSkill	= 0;
			pProject->activities[i].midSkill	= 1;
			pProject->activities[i].lowSkill	= 0;
		}
		else {
			pProject->activities[i].highSkill = 0;
			pProject->activities[i].midSkill = 0;
			pProject->activities[i].lowSkill	= 1;
		}
	}

	for (int i = 0; i < pProject->numActivities; ++i) {
		high += pProject->activities[i].highSkill* pProject->activities[i].duration;
		mid +=  pProject->activities[i].midSkill * pProject->activities[i].duration;
		low +=  pProject->activities[i].lowSkill * pProject->activities[i].duration;
	}

	return CalculateTotalLaborCost(high, mid, low);
}

// 등급별 투입 인력 계산 및 프로젝트의 수익 생성 함수
double CCreator::CalculateTotalLaborCost(int highCount, int midCount, int lowCount) {
	double highLaborCost	= CalculateLaborCost("H") * highCount;
	double midLaborCost	= CalculateLaborCost("M") * midCount;
	double lowLaborCost	= CalculateLaborCost("L") * lowCount;
	
	double totalLaborCost = highLaborCost + midLaborCost + lowLaborCost;
	return totalLaborCost;
}

// 등급별 투입 인력에 따른 수익 계산 함수
// 수정시 다음도 수정 필요 BOOL CCompany::Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad)
double CCreator::CalculateLaborCost(const std::string& grade) {
	double directLaborCost	= 0;
	double overheadCost	= 0;
	double technicalFee	= 0;
	double totalLaborCost	= 0;

	// 입력된 grade를 대문자로 변환
	char upperGrade = std::toupper(static_cast<unsigned char>(grade[0]));

	switch (upperGrade) {
	case 'H':
		directLaborCost = 50;
		break;
	case 'M':
		directLaborCost = 39;
		break;
	case 'L':
		directLaborCost = 25;
		break;
	default:
		AfxMessageBox(_T("잘못된 등급입니다. 'H', 'M', 'L' 중 하나를 입력하세요."), MB_OK | MB_ICONERROR);
		return 0; // 잘못된 입력 시 함수 종료
	}

	// 간접 : 직접 : 기술 = 6:2:2 = 10 ==> 소숫점이 나오지 않게 10배 키워서 계산한다.
	overheadCost = directLaborCost * 0.6; // 간접 비용 계산
	technicalFee = (directLaborCost + overheadCost) * 0.2; // 기술 비용 계산
	totalLaborCost = directLaborCost + overheadCost + technicalFee; // 총 인건비 계산

	return totalLaborCost;
}


// 대금 지급 조건 생성 함수
void CCreator::CalculatePaymentSchedule(PROJECT* pProject) {

	int paymentType;
	int paymentRatio;
	int totalPayments;

	pProject->firstPayMonth = 1;
	
	// 6주 이하의 짧은 프로젝트는 선금, 잔금만 있다.
	if (pProject->duration <= 6) {
		paymentType = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

		switch (paymentType) {
		case 1:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
			pProject->cashFlows[0] = 30;
			pProject->cashFlows[1] = 70;
			break;
		case 2:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.4);
			pProject->cashFlows[0] = 40;
			pProject->cashFlows[1] = 60;
			break;
		case 3:
			pProject->firstPay = (int)ceil((double)pProject->profit * 0.5);
			pProject->cashFlows[0] = 50;
			pProject->cashFlows[1] = 50;
			break;
		}

		pProject->secondPay = pProject->profit - pProject->firstPay;
		totalPayments = 2;
		pProject->secondPayMonth = pProject->duration;
	}

	// 7~12주 사이의 프로젝트는 3회에 걸처셔 받는다.
	else if (pProject->duration <= 12) {
		paymentType = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

		if (paymentType <= 3) {
			paymentRatio = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

			switch (paymentRatio) {
			case 1:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->cashFlows[0] = 30;
				pProject->cashFlows[1] = 70;
				break;
			case 2:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.4);
				pProject->cashFlows[0] = 40;
				pProject->cashFlows[1] = 60;
				break;
			case 3:
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.5);
				pProject->cashFlows[0] = 50;
				pProject->cashFlows[1] = 50;
				break;
			}

			pProject->secondPay = pProject->profit - pProject->firstPay;
			totalPayments = 2;
			pProject->secondPayMonth = pProject->duration;
		}
		else {
			paymentRatio = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

			if (paymentRatio <= 6) {
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->secondPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->cashFlows[0] = 30;
				pProject->cashFlows[1] = 30;
				pProject->cashFlows[2] = 40;
			}
			else {
				pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
				pProject->secondPay = (int)ceil((double)pProject->profit * 0.4);
				pProject->cashFlows[0] = 30;
				pProject->cashFlows[1] = 40;
				pProject->cashFlows[2] = 30;
			}

			pProject->finalPay = pProject->profit - pProject->firstPay - pProject->secondPay;
			totalPayments = 3;
			pProject->secondPayMonth = (int)ceil((double)pProject->duration / 2);
			pProject->finalPayMonth = pProject->duration;
		}
	}

	// 1년 이상의 프로젝트는 3회에 걸처서 받는다.
	else {
		pProject->firstPay = (int)ceil((double)pProject->profit * 0.3);
		pProject->secondPay = (int)ceil((double)pProject->profit * 0.4);
		pProject->finalPay = pProject->profit - pProject->firstPay - pProject->secondPay;

		pProject->cashFlows[0] = 30;
		pProject->cashFlows[1] = 40;
		pProject->cashFlows[2] = 30;

		totalPayments = 3;
		pProject->secondPayMonth = (int)ceil((double)pProject->duration / 2);
		pProject->finalPayMonth = pProject->duration;
	}
	pProject->nCashFlows = totalPayments;
}


void CCreator::Save(CString filename)
{
	ULONG ulTotalWritten = 0;
	FILE* fp = nullptr;
	if (!OpenFile(filename, _T("wb"), &fp)) return;
	
	
	SAVE_SIG sig;
	ulTotalWritten += fwrite(&sig, 1, sizeof(sig), fp);  // 파일 시작 부분에 시그니처 쓰기	
	ulTotalWritten += WriteDataWithHeader(fp, TYPE_ENVIRONMENT, &m_GlobalEnv, sizeof(GLOBAL_ENV));
	ulTotalWritten += WriteDataWithHeader(fp, TYPE_ACTIVITY, &m_ActType, sizeof(ALL_ACT_TYPE));
	ulTotalWritten += WriteDataWithHeader(fp, TYPE_PATTERN, &m_ActPattern, sizeof(ALL_ACTIVITY_PATTERN));

	// 모아서 적어야 한다.
	int size = m_orderTable.getCols() * m_orderTable.getRows();
	int* temp = new int[size];
	m_orderTable.copyToContinuousMemory(temp, size);

	ULONG orderTableSize = sizeof(int) * size;  // 바이트 단위로 크기 계산
	ulTotalWritten += WriteDataWithHeader(fp, TYPE_ORDER, temp, orderTableSize);
	delete[] temp;

	WriteProjet(fp);
	
	sig.totalLen = ulTotalWritten;
	fseek(fp, 0, SEEK_SET);  // 파일 포인터를 파일의 시작 위치로 이동	
	fwrite(&sig, 1, sizeof(sig), fp);  // 수정된 시그니처 다시 쓰기

	// 파일 닫기
	CloseFile(&fp);

}


void CCreator::WriteProjet(FILE* fp)
{
	SAVE_TL tl ;
	tl.length =  m_totalProjectNum;
	tl.type = TYPE_PROJECT;

	ULONG ulTemp = 0;
	ULONG ulWritten = 0;

	ulTemp = fwrite(&tl, sizeof(tl), 1, fp);  // 먼저 데이터 타입 및 길이 정보를 쓴다
	ulWritten += ulTemp * sizeof(tl);

	int test = sizeof(PROJECT);

	for(int i =0 ; i< m_totalProjectNum;i ++)
	{
		PROJECT* pProject; 
		pProject = m_pProjects+i;
		ulTemp = fwrite(pProject, sizeof(PROJECT), 1, fp);
		ulWritten += ulTemp;
	}

}

