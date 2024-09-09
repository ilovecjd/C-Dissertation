#pragma once

#include "globalenv.h"


class CProject;


class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(GLOBAL_ENV* pGlobalEnv, int Id, BOOL shouldLoad);
	//void PrintProjectInfo(SheetName sheet, CProject* pProject);

	/*void AllocateManageTable(MANAGE_TABLE* table, int size);
	void DeallocateManageTable(MANAGE_TABLE* table);*/

	void Load(CString fileName);
	BOOL LoadProjectsFromExcel();
	BOOL LoadProjects();	
	BOOL CreateProjects();

	BOOL Decision(int thisWeek);
	int CalculateFinalResult();
	void PrintProjectInfo(CXLEzAutomation* pXl, PROJECT* pProject);
	void PrintProjects();
	//void SaveProjectToAhn(const CString& filename);
	//void LoadProjectFromAhn();
		
	int* m_orderTable[2];

	int** m_doingHR;
	int** m_freeHR;
	int** m_totalHR;
	
	int** m_incomeTable;
	int** m_expensesTable;

	Dynamic2DArray m_doingTable;
	Dynamic2DArray m_doneTable;
	Dynamic2DArray m_defferTable;
	
	Dynamic2DArray m_debugInfo;
	
	//MANAGE_TABLE m_manageTable = {}; // NULL 로 초기화
	
	PROJECT* m_AllProjects;
	int m_totalProjectNum;

	CXLEzAutomation* m_pXl; // 엑셀을 다루기 위한 클래스	
	//COM_VAR com_var;
	void PrintDBTitle();

private:
	GLOBAL_ENV m_GlobalEnv;	
	ALL_ACT_TYPE	m_ActType;
	ALL_ACTIVITY_PATTERN m_ActPattern;

	void AllTableInit(int nWeeks);
	
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR(int thisWeek, CProject* project);
	void SelectNewProject(int thisWeek);
	void PrintDBData();

	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(CProject* project, int addWeek);
	void AddHR(int grade, int addWeek);
	void RemoveHR(int grade, int addWeek);

	void CCompany::ReadOrder(FILE* fp);
	void ReadProject(FILE* fp);
	
}; 

