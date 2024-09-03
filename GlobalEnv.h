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
//int** allocAndInit2DArray(int rows, int cols);
//int** del2DArray(int** array, int rows);
int PoissonRandom(double lambda);

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
} GLOBAL_ENV, *PGLOBAL_ENV;


// company 에서 사용

//////////////////////////////////////////////////////////////////////////
// activity의 타입에 대한 구조체
// activity_struct 시트의 cells(3,2) ~ cells(7,14)의 값으로 채워진다.
typedef struct {

	int occurrenceRate;     // 타입별 발생 확률 (%)
	int cumulativeRate;     // 누적 확률 (%)
	int minPeriod;          // 최소 기간
	int maxPeriod;          // 최대 기간
	int patternCount;       // 패턴 수

	// 반복되는 패턴 번호와 확률
	int patterns[4][2];     // 최대 5개의 패턴 번호와 확률을 저장하는 2차원 배열

} ACT_TYPE, *PACT_TYPE; // 구조체 이름과 포인터 타입 별칭

// activity의 속성에 대한 구조체와 정수 2차원 배열을 포함하는 유니온 정의
typedef union {
	ACT_TYPE actTypes[5];  // 5개의 타일 발생 데이터를 위한 구조체 배열
	int asIntArray[5][sizeof(ACT_TYPE) / sizeof(int)];  // 5개의 타일 데이터를 정수 배열로 접근 (2차원 배열)
} ALL_ACT_TYPE, *PALL_ACT_TYPE;
//////////////////////////////////////////////////////////////////////////


//////////////////////////////////////////////////////////////////////////
// activity_struct 시트의 cells(15,2) ~ cells(20,27)의 값으로 채워진다.
// 각 활동의 기간 비율과 인력 비율 패턴에 대한 구조체 정의
typedef struct {
	int minDurationRate;   // 최소 기간 비율 (%)
	int maxDurationRate;   // 최대 기간 비율 (%)
	int highHR;            // 고 인력 비율 (%)
	int mediumHR;          // 중 인력 비율 (%)
	int lowHR;             // 초 인력 비율 (%)
} ACT_PATTERN,*PACT_PATTERN;

// 모든 활동의 패턴을 포함하는 구조체 정의
typedef struct {
	int patternCount;    // 활동 패턴 갯수
	ACT_PATTERN patterns[5];  // 5개의 활동 패턴
} ALL_ACT_PATTERN,*PALL_ACT_PATTERN;

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

//typedef struct _MANAGE_TABLE{
//
//	// Order Table
//	//int* pWeeksNum;	// 주 (1,2,3,4,.... last week num)
//	//int* pSum;		// 누계
//	//int* pOrder;    // 발주
//	//
//	// HR Table 
//	//int* pDoingHR_H;    // 주별 투입 된 고급 인력 
//	//int* pDoingHR_M;    // 주별 투입 된 중급 인력 
//	//int* pDoingHR_L;    // 주별 투입 된 초급 인력 
//
//	//int* pFreeHR_H;    // 주별 여유 고급 인력 
//	//int* pFreeHR_M;    // 주별 여유 중급 인력 
//	//int* pFreeHR_L;    // 주별 여유 초급 인력 
//	//
//	//int* pTotalHR_H;    // 주별 보유 고급 인력 
//	//int* pTotalHR_M;    // 주별 보유 중급 인력 
//	//int* pTotalHR_L;    // 주별 보유 초급 인력 
//
//} MANAGE_TABLE, *PMANAGE_TABLE;


// Sheet enumeration for easy reference
enum SheetName {
	WS_NUM_PARAMETERS = 0,
	WS_NUM_DASHBOARD,
	WS_NUM_PROJECT,
	WS_NUM_ACTIVITY_STRUCT,
	WS_NUM_DEBUG_INFO,
	WS_NUM_SHEET_COUNT // Total number of sheets
};

extern LPOLESTR gSheetNames[WS_NUM_SHEET_COUNT];// = { L"parameters", L"dashboard", L"project", L"activity_struct", L"debuginfo" };

class GlobalEnv
{

private :

	bool bIsInit;						// init 함수가 한번만 호출되게	

public:
	GlobalEnv();
	~GlobalEnv();

	bool Init();	
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

