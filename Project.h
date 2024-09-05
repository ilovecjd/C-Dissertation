#pragma once



class CProject
{
public:
	CProject();
	~CProject();

	// song
public :
	
	PRJ_VAR prj_var;
	

	BOOL Init(int type, int ID, int ODate, ALL_ACT_TYPE* pActType, ALL_ACTIVITY_PATTERN* pActPattern);
	
private:	
	ALL_ACT_TYPE m_ActType;
	ALL_ACTIVITY_PATTERN m_ActPattern;
	BOOL CreateActivities();
	int CalculateHRAndProfit();
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule();

	int ZeroOrOneByProb(int probability);
	int RandomBetween(int low, int high);
};

