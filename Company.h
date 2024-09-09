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
	//void PrintProjectInfo(SheetName sheet, CProject* pProject);

	/*void AllocateManageTable(MANAGE_TABLE* table, int size);
	void DeallocateManageTable(MANAGE_TABLE* table);*/

	BOOL Decision(int thisWeek);
	int CalculateFinalResult();
	void PrintProjectInfo(CXLEzAutomation* pXl, PROJECT* pProject);
	void PrintProjects();
	//void SaveProjectToAhn(const CString& filename);
	//void LoadProjectFromAhn();
		
	GLOBAL_ENV m_GlobalEnv;
	

	void PrintDBTitle();

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
	Dynamic2DArray m_debugInfo;
	
	Dynamic2DArray m_incomeTable;
	Dynamic2DArray m_expensesTable;

	int m_lastDecisionWeek;
	
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR(int thisWeek, PROJECT* project);
	void SelectNewProject(int thisWeek);
	void PrintDBData();

	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(PROJECT* project, int addWeek);
	void AddHR(int grade, int addWeek);
	void RemoveHR(int grade, int addWeek);

	void CCompany::ReadOrder(FILE* fp);
	void ReadProject(FILE* fp);
	
}; 

