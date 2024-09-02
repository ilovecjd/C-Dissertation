#pragma once

#define MAX_N_CF  3
#define MAX_PRJ_TYPE 5
#define MAX_ACT  4

#define RND_HR_H  20
#define RND_HR_M  70


class CProject
{
public:
	CProject();
	~CProject();

	// song
public :
	
	
	// 프로젝트 속성
	ACTIVITY m_activities[MAX_ACT]; // 활동에 관한 정보를 기록하는 배열

	// Init 함수에서 초기화
	int m_category;		// 프로젝트 분류 (0: 외부 / 1: 내부)
	int m_ID;			// 프로젝트의 번호
	int m_orderDate;	// 발주일
	int m_startAvail;	// 시작 가능일
	int m_isStart;		// 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
	int m_experience;	// 경험 (0: 무경험 1: 유경험)
	int m_winProb;		// 성공 확률
	int m_nCashFlows;	// 비용 지급 횟수

	// CreateActivities 함수에서 초기화
	int m_endDate;		// 프로젝트 종료일
	int m_duration;		// 프로젝트의 총 기간

	// 
	int m_profit;	// 총 기대 수익 (HR 종속)

	// 현금 흐름
	int m_cashFlows[MAX_N_CF];	// 용역비를 받는 비율을 기록하는 배열
	int m_firstPay;		// 선금 액수
	int m_secondPay;		// 2차 지급 액수
	int m_finalPay;		// 3차 지급 액수
	int m_firstPayMonth;	// 선금 지급일
	int m_secondPayMonth;	// 2차 지급일
	int m_finalPayMonth;	// 3차 지급일

	// 활동
	int numActivities;          // 총 활동 수//    std::array<Activity, MAX_ACT> m_activities; // 활동에 관한 정보를 기록하는 배열

	// 참고용변수
	int m_projectType;		// activity_struct 시트의 어느 타입의 프로젝트인가
	int m_activityPattern;	// activity_struct 시트의 어느 패턴인가

	BOOL Init(int type, int ID, int ODate, PALL_ACT_TYPE pActType, PALL_ACTIVITY_PATTERN pActPattern);
	
private:	
	PALL_ACT_TYPE m_pActType;
	PALL_ACTIVITY_PATTERN m_pActPattern;
	BOOL CreateActivities();
	int CalculateHRAndProfit();
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule();

	int ZeroOrOneByProb(int probability);
	int RandomBetween(int low, int high);
};

