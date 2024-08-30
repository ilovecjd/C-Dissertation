#include "stdafx.h"
#include "Project.h"
#include <cstdlib>   // std::srand, std::rand
#include <ctime>     // std::time

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
	//CalculateHRandPfofit();
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