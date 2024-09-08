#pragma once

#include "globalenv.h"



// Order Tabe index for easy reference
enum OrderIndex{
	ORDER_SUM = 0,
	ORDER_ORD,
	ORDER_COUNT // Total number of OrderTable
};

// HR Tabe index for easy reference
enum HRIndex {
	HR_HIG = 0,
	HR_MID ,
	HR_LOW ,
	HR_COUNT // Total number of HR Table
};


class CProject;


class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad);
	void PrintProjectInfo(SheetName sheet, CProject* pProject);

	/*void AllocateManageTable(MANAGE_TABLE* table, int size);
	void DeallocateManageTable(MANAGE_TABLE* table);*/

	BOOL LoadProjectsFromExcel();
	BOOL CreateProjects();

	BOOL Decision(int thisWeek);
	int CalculateFinalResult();
	
	void SaveProjectToAhn(const CString& filename);
	void LoadProjectFromAhn();
		
	int** m_orderTable;

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
	
	CProject** m_AllProjects;

	COM_VAR com_var;

private:
	PGLOBAL_ENV m_pGlobalEnv;
	CXLEzAutomation* m_pXl; // 엑셀을 다루기 위한 클래스	
	ALL_ACT_TYPE*	m_pActType;
	ALL_ACTIVITY_PATTERN* m_pActPattern;

	void AllTableInit(int nWeeks);
	void PrintDBTitle();
	BOOL CheckLastWeek(int thisWeek);
	void SelectCandidates(int thisWeek);
	BOOL IsEnoughHR(int thisWeek, CProject* project);
	void SelectNewProject(int thisWeek);
	void PrintDBData();

	int m_candidateTable[MAX_CANDIDATES] = { 0, };
	void AddProjectEntry(CProject* project, int addWeek);
	void AddHR(int grade, int addWeek);
	void RemoveHR(int grade, int addWeek);
	
}; 

