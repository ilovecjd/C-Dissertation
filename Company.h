#pragma once

class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad);
private:
	PGLOBAL_ENV m_pGlobalEnv;
	CXLEzAutomation* m_pXl; // 엑셀을 다루기 위한 클래스
	PALL_ACT_PATTERN m_pActPattern;
	PACT_TYPE	m_pActType;
	
}; 

