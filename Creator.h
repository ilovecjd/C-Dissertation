#pragma once

#include "globalenv.h"

class CCreator
{
public:
	CCreator();
	~CCreator();

	// song
public :
	
	int m_totalProjectNum;
	Dynamic2DArray m_orderTable;
	//PRJ_VAR prj_var;
	
	BOOL Init(GLOBAL_ENV* pGlobalEnv, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern);
	//BOOL Init(int type, int ID, int ODate, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern);
	void Save(CString filename);
	void Load(CString filename);

private:	
	GLOBAL_ENV m_GlobalEnv;
	ALL_ACT_TYPE m_ActType;
	ALL_ACTIVITY_PATTERN m_ActPattern;
	PROJECT* m_pProjects;


	
	int CreateOrderTable();
	int CreateProjects();
	BOOL CreateActivities(PROJECT* pProject);
	int CalculateHRAndProfit(PROJECT* pProject);
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule(PROJECT* pProject);

	int ZeroOrOneByProb(int probability);
	int RandomBetween(int low, int high);

	void WriteProjet(FILE* fp);
};

