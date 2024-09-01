#pragma once

class CProject;

class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad);
	void PrintProjectInfo(CProject* pProject);

	void AllocateManageTable(MANAGE_TABLE* table, int size);
	void DeallocateManageTable(MANAGE_TABLE* table);
	void testFunction();

	MANAGE_TABLE m_manageTable = {}; // NULL 로 초기화
	int		m_totalProjectNum;
	CProject** m_ProjectTable;

private:
	PGLOBAL_ENV m_pGlobalEnv;
	CXLEzAutomation* m_pXl; // 엑셀을 다루기 위한 클래스	
	PALL_ACT_TYPE	m_pActType;
	PALL_ACTIVITY_PATTERN m_pActPattern;
	
}; 

