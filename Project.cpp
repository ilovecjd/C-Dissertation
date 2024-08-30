#include "stdafx.h"
#include "Project.h"
#include <cstdlib>   // std::srand, std::rand
#include <ctime>     // std::time
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

BOOL CProject::Init(PALL_ACT_TYPE pActType, PALL_ACTIVITY_PATTERN pActPattern)
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

	std::memcpy(m_pActType, pActType, sizeof(ALL_ACT_TYPE));
	std::memcpy(m_pActPattern, pActPattern, sizeof(ALL_ACTIVITY_PATTERN));
	

	CreateActivities();
	CalculateHRAndProfit();
	//CalculatePaymentSchedule();
	return TRUE;
}



// 랜덤 숫자 생성 함수
int RandomBetween(int low, int high) {
	return low + std::rand() % (high - low + 1);
}

BOOL CProject::CreateActivities() {
	//song 사용하지 않는 멤버 변수와 지역 변수들 삭제 하자
	std::srand(static_cast<unsigned int>(std::time(0))); // 랜덤 시드 초기화

	int prjType = 0;
	int patternType = 0;
	int index;
	int probability;
	int Lb = 0;
	int UB = 0;
	int maxLoop;
	int totalDuration;
	int tempDuration;

	probability = std::rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
	maxLoop = MAX_PRJ_TYPE;

	////////////////////////////////////////////
	// 프로젝트 타입관련 정보
	// 프로젝트 타입 결정
	for (index = 0; index < maxLoop; ++index) {
		UB += m_pActType->asIntArray[index][0];//발생 확률

		if (Lb <= probability && probability < UB) {
			prjType = index;
			break;
		}

		Lb = UB;
	}

	
	Lb = m_pActType->asIntArray[prjType][2];//최소기간
	UB = m_pActType->asIntArray[prjType][3];//최대기간
	totalDuration = RandomBetween(Lb, UB);
	projectDuration = totalDuration;
	endDate = possiblestartDate + totalDuration - 1;// song??

	Lb = 0;
	UB = 0;
	maxLoop = m_pActType->asIntArray[prjType][4];//패턴수

	// 패턴 타입 결정
	for (index = 0; index < maxLoop; ++index) {
		UB += m_pActType->asIntArray[prjType][6 + ((index) * 2)];//1번패턴 확률부터

		if (Lb <= probability && probability < UB) {
			patternType = m_pActType->asIntArray[prjType][5 + ((index) * 2)];//1번패턴 패턴번호부터
			break;
		}
		Lb = UB;
	}


	//////////////////////////////////////////////////////////////////
	//프로젝트 패턴 관련 정보
	Lb = 0;
	UB = 0;
	maxLoop = m_pActPattern->asIntArray[patternType-1][0];//활동수 !!! -1 에 주의
	numActivities = maxLoop;

	// 활동 생성
	for (index = 0; index < maxLoop; ++index) {
		Lb += m_pActPattern->asIntArray[patternType-1][1 + ((index) * 5)];// !!! -1 에 주의
		UB += m_pActPattern->asIntArray[patternType-1][2 + ((index) * 5)];// !!! -1 에 주의
		probability = RandomBetween(Lb, UB);
		tempDuration = totalDuration * probability / 100;

		if (tempDuration == 0) {
			tempDuration = 1;
		}

		if (index == 0) {
			m_activities[index].duration = tempDuration;
			m_activities[index].startDate = possiblestartDate;
			m_activities[index].endDate = possiblestartDate - 1 + tempDuration;
		}
		else if (index == 1) {
			m_activities[index].duration = totalDuration - m_activities[0].duration;
			m_activities[index].startDate = m_activities[0].endDate + 1;
			m_activities[index].endDate = possiblestartDate - 1 + totalDuration;
		}
		else if (index == 2) {
			m_activities[index].duration = tempDuration;
			m_activities[index].startDate = possiblestartDate - 1 + totalDuration - tempDuration + 1;
			m_activities[index].endDate = possiblestartDate - 1 + totalDuration;
		}
		else {
			m_activities[index].duration = tempDuration;
			m_activities[index].startDate = m_activities[2].startDate - tempDuration;
			m_activities[index].endDate = m_activities[2].startDate - 1;
		}
	}
	return TRUE;
}



// 활동별 투입 인력 생성 및 프로젝트 전체 기대 수익 계산 함수
void CProject::CalculateHRAndProfit() {
	int high = 0, mid = 0, low = 0;

	std::srand(static_cast<unsigned int>(std::time(0))); // 랜덤 시드 초기화

	for (int index = 0; index < numActivities; ++index) {
		int j = std::rand() % 100; // 0부터 99 사이의 랜덤 정수 생성
		if (0 < j && j <= RND_HR_H) {
			m_activities[index].highSkill = 1;
			m_activities[index].midSkill = 0;
			m_activities[index].lowSkill = 0;
		}
		else if (RND_HR_H < j && j <= RND_HR_M) {
			m_activities[index].highSkill = 0;
			m_activities[index].midSkill = 1;
			m_activities[index].lowSkill = 0;
		}
		else {
			m_activities[index].highSkill = 0;
			m_activities[index].midSkill = 0;
			m_activities[index].lowSkill = 1;
		}
	}

	for (int index = 0; index < numActivities; ++index) {
		high += m_activities[index].highSkill * m_activities[index].duration;
		mid += m_activities[index].midSkill * m_activities[index].duration;
		low += m_activities[index].lowSkill * m_activities[index].duration;
	}

	profit = CalculateTotalLaborCost(high, mid, low);
}

// 등급별 투입 인력 계산 및 프로젝트의 수익 생성 함수
double CProject::CalculateTotalLaborCost(int highCount, int midCount, int lowCount) {
	double highLaborCost = CalculateLaborCost("H") * highCount;
	double midLaborCost = CalculateLaborCost("M") * midCount;
	double lowLaborCost = CalculateLaborCost("L") * lowCount;

	double totalLaborCost = highLaborCost + midLaborCost + lowLaborCost;
	return totalLaborCost;
}

// 등급별 투입 인력에 따른 수익 계산 함수
double CProject::CalculateLaborCost(const std::string& grade) {
	double directLaborCost = 0.0;
	double overheadCost = 0.0;
	double technicalFee = 0.0;
	double totalLaborCost = 0.0;

	// 입력된 grade를 대문자로 변환
	char upperGrade = std::toupper(static_cast<unsigned char>(grade[0]));

	switch (upperGrade) {
	case 'H':
		directLaborCost = 50.0;
		break;
	case 'M':
		directLaborCost = 39.0;
		break;
	case 'L':
		directLaborCost = 25.0;
		break;
	default:
		AfxMessageBox(_T("잘못된 등급입니다. 'H', 'M', 'L' 중 하나를 입력하세요."), MB_OK | MB_ICONERROR);
		return 0.0; // 잘못된 입력 시 함수 종료
	}

	overheadCost = directLaborCost * 0.6; // 간접 비용 계산
	technicalFee = (directLaborCost + overheadCost) * 0.2; // 기술 비용 계산
	totalLaborCost = directLaborCost + overheadCost + technicalFee; // 총 인건비 계산

	return totalLaborCost;
}