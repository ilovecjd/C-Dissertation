#pragma once



class CProject
{
public:
	CProject();
	~CProject();

	// song
public :
	
	PRJ_VAR prj_var;
	

	BOOL Init(int type, int ID, int ODate, PALL_ACT_TYPE pActType, PALL_ACTIVITY_PATTERN pActPattern);
	
private:	
	PALL_ACT_TYPE m_pActType;
	PALL_ACTIVITY_PATTERN m_pActPattern;
	BOOL CreateActivities();
	int CalculateHRAndProfit();
	double CalculateTotalLaborCost(int highCount, int midCount, int lowCount);
	double CalculateLaborCost(const std::string& grade);
	void CalculatePaymentSchedule();

	int ZeroOrOneByProb(int probability);
	int RandomBetween(int low, int high);
};

