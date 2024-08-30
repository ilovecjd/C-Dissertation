﻿#pragma once

class CProject;

class CCompany
{
public:
	CCompany();
	~CCompany();
	BOOL Init(PGLOBAL_ENV pGlobalEnv, int Id, BOOL shouldLoad);
	void PrintProjectInfo(CProject* pProject);
	void testFunction();
private:
	PGLOBAL_ENV m_pGlobalEnv;
	CXLEzAutomation* m_pXl; // 엑셀을 다루기 위한 클래스	
	PALL_ACT_TYPE	m_pActType;
	PALL_ACTIVITY_PATTERN m_pActPattern;
	
}; 

