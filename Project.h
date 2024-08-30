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
	BOOL Init(PALL_ACT_TYPE pActType, PALL_ACTIVITY_PATTERN pActPattern);

// 프로젝트 속성
	ACTIVITY m_activities[MAX_ACT]; // 활동에 관한 정보를 기록하는 배열

    int projectType;            // 프로젝트 타입 (0: 외부 / 1: 내부)
    int projectNum;             // 프로젝트의 번호

    int orderDate;              // 발주일
    int possiblestartDate;      // 시작 가능일
    int endDate;                // 프로젝트 종료일
    int projectDuration;        // 프로젝트의 총 기간
    int isStart;                // 진행 여부 (0: 미진행, 나머지: 진행시작한 주)
    double profit;              // 총 기대 수익 (HR 종속)
    int experience;             // 경험 (0: 무경험 1: 유경험)
    int successProbability;     // 성공 확률

    // 현금 흐름
    int numCashFlows;           // 비용 지급 횟수
    std::array<int, MAX_N_CF> m_cashFlows; // 용역비를 받는 비율을 기록하는 배열
    long firstPayment;          // 선금 액수
    long middlePayment;         // 2차 지급 액수
    long finalPayment;          // 3차 지급 액수
    int firstPaymentMonth;      // 선금 지급일
    int middlePaymentMonth;     // 2차 지급일
    int finalPaymentMonth;      // 3차 지급일

    // 활동
    int numActivities;          // 총 활동 수
//    std::array<Activity, MAX_ACT> m_activities; // 활동에 관한 정보를 기록하는 배열
//    std::vector<std::variant<int, double, std::string>> activityAttribute; // 다양한 타입을 가질 수 있는 배열
  //  std::vector<std::variant<int, double, std::string>> activityPattern;   // 다양한 타입을 가질 수 있는 배열

private:	
	PALL_ACT_TYPE m_pActType;
	PALL_ACTIVITY_PATTERN m_pActPattern;
	BOOL CreateActivities();
	void CalculateHRAndProfit();
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
};

