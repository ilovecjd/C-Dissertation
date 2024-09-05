#include "stdafx.h"
#include "GlobalEnv.h"
#include "Project.h"
#include <cctype>   // toupper 함수를 사용하기 위해 필요

CProject::CProject()
{
	m_pActType = new ALL_ACT_TYPE;
	m_pActPattern = new ALL_ACTIVITY_PATTERN;	
}

CProject::~CProject()
{	
	delete m_pActType;
	delete m_pActPattern;
}

BOOL CProject::Init(int type, int ID, int ODate, PALL_ACT_TYPE pActType, PALL_ACTIVITY_PATTERN pActPattern)
{
	// song 한번만 실행되게 하는 코드 추가 필요
	if (m_pActType == nullptr || pActType == nullptr) {
		MessageBox(NULL, _T("pActType is NULL."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}

	if (m_pActPattern == nullptr || pActPattern == nullptr) {
		MessageBox(NULL, _T("pActPattern is NULL."), _T("Error"), MB_OK | MB_ICONERROR);
		return FALSE;
	}
	
	// 시작 가능일 계산
	std::memcpy(m_pActType,		pActType,	 sizeof(ALL_ACT_TYPE));
	std::memcpy(m_pActPattern,	pActPattern, sizeof(ALL_ACTIVITY_PATTERN));
	
	prj_var.m_category		= type;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	prj_var.m_ID			= ID;		// 프로젝트의 번호	
	prj_var.m_orderDate		= ODate;	// 발주일
	prj_var.m_startAvail	= ODate + (rand() % 4);  // // 시작 가능일 ( 0에서 3 사이의 정수 난수 생성)
	prj_var.m_isStart		= 0;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
	prj_var.m_experience	= ZeroOrOneByProb(95);	// 경험 (0: 무경험 1: 유경험)
	prj_var.m_winProb		= 100;		// 성공 확률 song ==> 추후 사용시 생성 방법을 결정한다. 현재는 100%
	prj_var.m_nCashFlows	= MAX_N_CF;	// 비용 지급 횟수(규모에 따라 변경 가능)

	CreateActivities();					//m_activities[MAX_ACT] 계산
	prj_var.m_profit = CalculateHRAndProfit();
	CalculatePaymentSchedule();			//m_cashFlows[MAX_N_CF] 계산
	return TRUE;
}


// 확률에 따라서 0 또는 1 생성
int CProject::ZeroOrOneByProb(int probability)
{
	double randomProb = (double)rand() / RAND_MAX;
	return (randomProb <= (double)probability / 100) ? 1 : 0;
}

// 랜덤 숫자 생성 함수
int CProject::RandomBetween(int low, int high) {
	return low + rand() % (high - low + 1);
}

BOOL CProject::CreateActivities() {
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
		UB += m_pActType->asIntArray[i][0];	// 엑셀 2열의 "발생 확률"

		if (Lb <= probability && probability < UB) {
			prj_var.m_projectType = i;
			break;
		}

		Lb = UB;
	}
	
	Lb = m_pActType->asIntArray[prj_var.m_projectType][2];	// 엑셀 4열의 "최소기간"
	UB = m_pActType->asIntArray[prj_var.m_projectType][3];	// 엑셀 5열의 "최대기간"

	totalDuration = RandomBetween(Lb, UB);
	prj_var.m_duration = totalDuration;
	prj_var.m_endDate = prj_var.m_startAvail + totalDuration - 1;// song??

	Lb = 0;
	UB = 0;
	maxLoop = m_pActType->asIntArray[prj_var.m_projectType][4];//패턴수

	// 패턴 타입 결정
	for (i = 0; i < maxLoop; ++i) {
		UB += m_pActType->asIntArray[prj_var.m_projectType][6 + ((i) * 2)];//1번패턴 확률부터

		if (Lb <= probability && probability < UB) {
			prj_var.m_activityPattern = m_pActType->asIntArray[prj_var.m_projectType][5 + ((i) * 2)];//1번패턴 패턴번호부터
			break;
		}
		Lb = UB;
	}

	//////////////////////////////////////////////////////////////////
	//프로젝트 패턴 관련 정보
	Lb = 0;
	UB = 0;
	maxLoop = m_pActPattern->asIntArray[prj_var.m_activityPattern-1][0];//활동수 !!! -1 에 주의
	prj_var.numActivities = maxLoop;

	// 활동 생성
	for (i = 0; i < maxLoop; ++i) {
		Lb += m_pActPattern->asIntArray[prj_var.m_activityPattern-1][1 + ((i) * 5)];// !!! -1 에 주의
		UB += m_pActPattern->asIntArray[prj_var.m_activityPattern-1][2 + ((i) * 5)];// !!! -1 에 주의
		probability = RandomBetween(Lb, UB);
		tempDuration = totalDuration * probability / 100;

		if (tempDuration == 0) {
			tempDuration = 1;
		}

		if (i == 0) {
			prj_var.m_activities[i].duration = tempDuration;
			prj_var.m_activities[i].startDate = prj_var.m_startAvail;
			prj_var.m_activities[i].endDate = prj_var.m_startAvail - 1 + tempDuration;
		}
		else if (i == 1) {
			prj_var.m_activities[i].duration = totalDuration - prj_var.m_activities[0].duration;
			prj_var.m_activities[i].startDate = prj_var.m_activities[0].endDate + 1;
			prj_var.m_activities[i].endDate = prj_var.m_startAvail - 1 + totalDuration;
		}
		else if (i == 2) {
			prj_var.m_activities[i].duration = tempDuration;
			prj_var.m_activities[i].startDate = prj_var.m_startAvail - 1 + totalDuration - tempDuration + 1;
			prj_var.m_activities[i].endDate = prj_var.m_startAvail - 1 + totalDuration;
		}
		else {
			prj_var.m_activities[i].duration = tempDuration;
			prj_var.m_activities[i].startDate = prj_var.m_activities[2].startDate - tempDuration;
			prj_var.m_activities[i].endDate = prj_var.m_activities[2].startDate - 1;
		}
	}
	return TRUE;
}

// 활동별 투입 인력 생성 및 프로젝트 전체 기대 수익 계산 함수
int CProject::CalculateHRAndProfit() {
	int high = 0, mid = 0, low = 0;

	for (int i = 0; i < prj_var.numActivities; ++i) {
		int j = rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
		if (0 < j && j <= RND_HR_H) {
			prj_var.m_activities[i].highSkill	= 1;
			prj_var.m_activities[i].midSkill	= 0;
			prj_var.m_activities[i].lowSkill	= 0;
		}
		else if (RND_HR_H < j && j <= RND_HR_M) {
			prj_var.m_activities[i].highSkill	= 0;
			prj_var.m_activities[i].midSkill	= 1;
			prj_var.m_activities[i].lowSkill	= 0;
		}
		else {
			prj_var.m_activities[i].highSkill	= 0;
			prj_var.m_activities[i].midSkill	= 0;
			prj_var.m_activities[i].lowSkill	= 1;
		}
	}

	for (int i = 0; i < prj_var.numActivities; ++i) {
		high += prj_var.m_activities[i].highSkill* prj_var.m_activities[i].duration;
		mid +=  prj_var.m_activities[i].midSkill * prj_var.m_activities[i].duration;
		low +=  prj_var.m_activities[i].lowSkill * prj_var.m_activities[i].duration;
	}

	return CalculateTotalLaborCost(high, mid, low);
}

// 등급별 투입 인력 계산 및 프로젝트의 수익 생성 함수
double CProject::CalculateTotalLaborCost(int highCount, int midCount, int lowCount) {
	double highLaborCost	= CalculateLaborCost("H") * highCount;
	double midLaborCost	= CalculateLaborCost("M") * midCount;
	double lowLaborCost	= CalculateLaborCost("L") * lowCount;
	
	double totalLaborCost = highLaborCost + midLaborCost + lowLaborCost;
	return totalLaborCost;
}

// 등급별 투입 인력에 따른 수익 계산 함수
// 수정시 다음도 수정 필요 BOOL CCompany::Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad)
double CProject::CalculateLaborCost(const std::string& grade) {
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
void CProject::CalculatePaymentSchedule() {

	int paymentType;
	int paymentRatio;
	int totalPayments;

	prj_var.m_firstPayMonth = 1;
	
	// 6주 이하의 짧은 프로젝트는 선금, 잔금만 있다.
	if (prj_var.m_duration <= 6) {
		paymentType = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

		switch (paymentType) {
		case 1:
			prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.3);
			prj_var.m_cashFlows[0] = 30;
			prj_var.m_cashFlows[1] = 70;
			break;
		case 2:
			prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.4);
			prj_var.m_cashFlows[0] = 40;
			prj_var.m_cashFlows[1] = 60;
			break;
		case 3:
			prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.5);
			prj_var.m_cashFlows[0] = 50;
			prj_var.m_cashFlows[1] = 50;
			break;
		}

		prj_var.m_secondPay = prj_var.m_profit - prj_var.m_firstPay;
		totalPayments = 2;
		prj_var.m_secondPayMonth = prj_var.m_duration;
	}

	// 7~12주 사이의 프로젝트는 3회에 걸처셔 받는다.
	else if (prj_var.m_duration <= 12) {
		paymentType = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

		if (paymentType <= 3) {
			paymentRatio = rand() % 3 + 1;  // 1에서 3 사이의 난수 생성

			switch (paymentRatio) {
			case 1:
				prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.3);
				prj_var.m_cashFlows[0] = 30;
				prj_var.m_cashFlows[1] = 70;
				break;
			case 2:
				prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.4);
				prj_var.m_cashFlows[0] = 40;
				prj_var.m_cashFlows[1] = 60;
				break;
			case 3:
				prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.5);
				prj_var.m_cashFlows[0] = 50;
				prj_var.m_cashFlows[1] = 50;
				break;
			}

			prj_var.m_secondPay = prj_var.m_profit - prj_var.m_firstPay;
			totalPayments = 2;
			prj_var.m_secondPayMonth = prj_var.m_duration;
		}
		else {
			paymentRatio = rand() % 10 + 1;  // 1에서 10 사이의 난수 생성

			if (paymentRatio <= 6) {
				prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.3);
				prj_var.m_secondPay = (int)ceil((double)prj_var.m_profit * 0.3);
				prj_var.m_cashFlows[0] = 30;
				prj_var.m_cashFlows[1] = 30;
				prj_var.m_cashFlows[2] = 40;
			}
			else {
				prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.3);
				prj_var.m_secondPay = (int)ceil((double)prj_var.m_profit * 0.4);
				prj_var.m_cashFlows[0] = 30;
				prj_var.m_cashFlows[1] = 40;
				prj_var.m_cashFlows[2] = 30;
			}

			prj_var.m_finalPay = prj_var.m_profit - prj_var.m_firstPay - prj_var.m_secondPay;
			totalPayments = 3;
			prj_var.m_secondPayMonth = (int)ceil((double)prj_var.m_duration / 2);
			prj_var.m_finalPayMonth = prj_var.m_duration;
		}
	}

	// 1년 이상의 프로젝트는 3회에 걸처서 받는다.
	else {
		prj_var.m_firstPay = (int)ceil((double)prj_var.m_profit * 0.3);
		prj_var.m_secondPay = (int)ceil((double)prj_var.m_profit * 0.4);
		prj_var.m_finalPay = prj_var.m_profit - prj_var.m_firstPay - prj_var.m_secondPay;

		prj_var.m_cashFlows[0] = 30;
		prj_var.m_cashFlows[1] = 40;
		prj_var.m_cashFlows[2] = 30;

		totalPayments = 3;
		prj_var.m_secondPayMonth = (int)ceil((double)prj_var.m_duration / 2);
		prj_var.m_finalPayMonth = prj_var.m_duration;
	}

	prj_var.m_nCashFlows = totalPayments;
}
