#pragma once

#include "globalenv.h"


class CProject;


class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(CString fileName);
	void ReInit();
	void ClearMemory();
	void PrintResult(CString fileName); // 시뮬레이션 결과를 엑셀에 출력한다.
	void SaveResult(CString fileName); // 시뮬레이션 결과를 파일에 저장한다.
	void LoadResult(CString fileName); // 시뮬레이션 결과를 파일에서 읽어온다.
	//void PrintProjectInfo(SheetName sheet, CProject* pProject);
	
	BOOL Decision(int thisWeek);
	int CalculateFinalResult();
	void PrintProjectInfo(CXLEzAutomation* pXl, PROJECT* pProject);
	void PrintProjects(CXLEzAutomation* pxl);
	//void SaveProjectToAhn(const CString& filename);
	//void LoadProjectFromAhn();
		
	GLOBAL_ENV m_GlobalEnv;
	int m_lastDecisionWeek;
	

	void PrintDBTitle(CXLEzAutomation* pXl);

private:
	// 초기화 필요한 변수들
	int m_totalProjectNum;

	ALL_ACT_TYPE	m_ActType;
	ALL_ACTIVITY_PATTERN m_ActPattern;

	PROJECT* m_AllProjects = NULL;	
	CXLEzAutomation* m_pXl = NULL; // 엑셀을 다루기 위한 클래스	

	int* m_orderTable[2] = {NULL,NULL};

	Dynamic2DArray m_doingHR;
	Dynamic2DArray m_freeHR;
	Dynamic2DArray m_totalHR;

	Dynamic2DArray m_doingTable;
	Dynamic2DArray m_doneTable;
	Dynamic2DArray m_defferTable;
		
	Dynamic2DArray m_incomeTable;
	Dynamic2DArray m_expensesTable;

	
	
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR(int thisWeek, PROJECT* project);
	void SelectNewProject(int thisWeek);
	void PrintDBData(CXLEzAutomation* pXl);

	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(PROJECT* project, int addWeek);
	void AddHR(int grade, int addWeek);
	void RemoveHR(int grade, int addWeek);

	void CCompany::ReadOrder(FILE* fp);
	void ReadProject(FILE* fp);
	
}; 

