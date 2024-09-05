#pragma once

#include <iostream>
#include <vector>
#include <algorithm>

// 엑셀 파일 관리
// 전역 환경변수 관리

#define __STR_DATA_FILE		"data.xlsm"
#define __STR_RUN_LOG_FILE	"run_log.txt"
#define __STR_START_EXCEL
#define __STR_END_EXCEL
#define __NUN_OF_COMPANY	1 // 


#define MAX_CANDIDATES 50
#define ADD_HR_SIZE 80

int PoissonRandom(double lambda);

// project 에서 사용
#define MAX_N_CF  3
#define MAX_PRJ_TYPE 5
#define MAX_ACT  4

#define RND_HR_H  20
#define RND_HR_M  70


// company 에서 사용

//////////////////////////////////////////////////////////////////////////
// activity의 타입에 대한 구조체
// activity_struct 시트의 cells(3,2) ~ cells(7,14)의 값으로 채워진다.
struct ACT_TYPE {

	int occurrenceRate;     // 타입별 발생 확률 (%)
	int cumulativeRate;     // 누적 확률 (%)
	int minPeriod;          // 최소 기간
	int maxPeriod;          // 최대 기간
	int patternCount;       // 패턴 수

	// 반복되는 패턴 번호와 확률
	int patterns[4][2];     // 최대 5개의 패턴 번호와 확률을 저장하는 2차원 배열
}; 

// activity의 속성에 대한 구조체와 정수 2차원 배열을 포함하는 유니온 정의
union ALL_ACT_TYPE {
	ACT_TYPE actTypes[5];  // 5개의 타일 발생 데이터를 위한 구조체 배열
	int asIntArray[5][sizeof(ACT_TYPE) / sizeof(int)];  // 5개의 타일 데이터를 정수 배열로 접근 (2차원 배열)
} ;
//////////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////////
// activity_struct 시트의 cells(15,2) ~ cells(20,27)의 값으로 채워진다.
// 각 활동의 기간 비율과 인력 비율 패턴에 대한 구조체 정의
struct ACT_PATTERN {
	int minDurationRate;   // 최소 기간 비율 (%)
	int maxDurationRate;   // 최대 기간 비율 (%)
	int highHR;            // 고 인력 비율 (%)
	int mediumHR;          // 중 인력 비율 (%)
	int lowHR;             // 초 인력 비율 (%)
};

// 모든 활동의 패턴을 포함하는 구조체 정의
struct ALL_ACT_PATTERN {
	int patternCount;    // 활동 패턴 갯수
	ACT_PATTERN patterns[5];  // 5개의 활동 패턴
} ;

// 활동 패턴을 정수 2차원 배열로도 접근할 수 있는 유니온 정의
typedef union {
	ALL_ACT_PATTERN pattern[6];  // 6개의 활동 패턴을위한 구조체 배열
	int asIntArray[6][sizeof(ALL_ACT_PATTERN) / sizeof(int)];  // 6개의 활동 데이터를 정수 배열로 접근 (2차원 배열)
} ALL_ACTIVITY_PATTERN, *PALL_ACTIVITY_PATTERN;
//////////////////////////////////////////////////////////////////////////


// Company 에서 사용
typedef struct _ACTIVITY {
	int activityType;  // 활동 유형
    int duration;      // 활동 기간
    int startDate;     // 시작 날짜
    int endDate;       // 종료 날짜
    int highSkill;     // 높은 기술 수준 인력 수
    int midSkill;      // 중간 기술 수준 인력 수
    int lowSkill;      // 낮은 기술 수준 인력 수
} ACTIVITY, *PACTIVITY;

// Sheet enumeration for easy reference
enum SheetName {
	WS_NUM_PARAMETERS = 0,
	WS_NUM_DASHBOARD,
	WS_NUM_PROJECT,
	WS_NUM_ACTIVITY_STRUCT,
	WS_NUM_DEBUG_INFO,
	WS_NUM_SHEET_COUNT // Total number of sheets
};



typedef struct {
	int		SimulationWeeks;
	int		Hr_TableSize;		//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	double	WeeklyProb;
	int		Hr_Init_H;
	int		Hr_Init_M;
	int		Hr_Init_L;
	int		Hr_LeadTime;
	int		Cash_Init;
	int		ProblemCnt;
	int		status;				// 프로그램의 동작 상태. 0:프로젝트 미생성, 1:프로젝트 생성,
								//////////////////////////////////////
								// 엑셀파일 오픈을 막자
	ALL_ACT_TYPE* ActType;
	ALL_ACTIVITY_PATTERN* ActPattern;

	// 정책을 설정한다.
	double	ExpenseRate;	// 비용계산에 사용되는 제경비 비율
	//double	profitRate;		// 프로젝트 총비용 계산에 사용되는 제경비 비율

	int		selectOrder;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로

	int		recruit;		// 충원에 필요한 운영비 (몇주분량인가?)
	int		layoff;			// 감원에 필요한 운영비 (몇주분량인가?)

} GLOBAL_ENV, *PGLOBAL_ENV;


struct COM_VAR{
	GLOBAL_ENV ge;
	int				m_ID;
	int m_lastDecisionWeek = 0;
	int		m_totalProjectNum;

};

struct PRJ_VAR {
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
};


class Dynamic2DArray {
private:
	std::vector<std::vector<int>> data;

public:
	Dynamic2DArray() {}

	Dynamic2DArray(const Dynamic2DArray& other) : data(other.data) {} // 복사 생성자

	Dynamic2DArray& operator=(const Dynamic2DArray& other) { // 할당 연산자
		if (this != &other) {
			data = other.data;
		}
		return *this;
	}

	class Proxy {
	private:
		Dynamic2DArray& array;
		int rowIndex;

	public:
		Proxy(Dynamic2DArray& arr, int index) : array(arr), rowIndex(index) {}

		int& operator[](int colIndex) {
			if (colIndex >= array.data[rowIndex].size()) {
				array.data[rowIndex].resize(colIndex + 1, 0);
			}
			return array.data[rowIndex][colIndex];
		}
	};

	Proxy operator[](int rowIndex) {
		if (rowIndex >= data.size()) {
			data.resize(rowIndex + 1);
		}
		return Proxy(*this, rowIndex);
	}

	void Resize(int x, int y) {
		data.resize(x);
		for (int i = 0; i < x; ++i) {
			data[i].resize(y, 0);
		}
	}

	void copyFromContinuousMemory(int* src, int rows, int cols) {
		Resize(rows, cols);
		for (int i = 0; i < rows; ++i) {
			for (int j = 0; j < cols; ++j) {
				data[i][j] = src[i * cols + j];
			}
		}
	}

	// Copy data to an external continuous memory block
	void copyToContinuousMemory(int* dest, int maxElements) {
		int index = 0;
		for (int i = 0; i < data.size() && index < maxElements; ++i) {
			for (int j = 0; j < data[i].size() && index < maxElements; ++j) {
				dest[index++] = data[i][j];
			}
		}
	}

	int getRows() const {
		return data.size();
	}

	int getCols() const {
		if (!data.empty()) return data[0].size();
		return 0;
	}
};


int** Newallocate2DArray(int rows, int cols);
void Newinitialize2DArray(int** array, int rows, int cols, int value);
void Newdeallocate2DArray(int** array, int rows);
void Newcopy2DArray(int** source, int** destination, int rows, int cols);

//	int** array = allocate2DArray(rows, cols);

