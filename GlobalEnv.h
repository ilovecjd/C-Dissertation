#pragma once

// 엑셀 파일 관리
// 전역 환경변수 관리

#define __STR_DATA_FILE		"data.xlsm"
#define __STR_RUN_LOG_FILE	"run_log.txt"
#define __STR_START_EXCEL
#define __STR_END_EXCEL
#define __NUN_OF_COMPANY	1 // 

#define __PARAMETERS_SHEET_NAME = "parameters"
#define __DBOARD_SHEET_NAME		= "dashboard"
#define __PROJECT_SHEET_NAME	= "project"
#define __ACTIVITY_SHEET_NAME	= "activity_struct"
#define __DEBUGINFO_SHEET_NAME	= "debuginfo"


class GlobalEnv
{

private :

	bool bIsInit;						// init 함수가 한번만 호출되게	

public:
	GlobalEnv();
	~GlobalEnv();

	
	bool Init();
	void LoadEnvFromExcel();

	xlnt::workbook xlWb;				// workbook
	xlnt::worksheet WsParameters;		// Parameters 시트 객체
	xlnt::worksheet WsDashboard;		// Dashboard 시트 객체
	xlnt::worksheet WsProject;			// Project 시트 객체
	xlnt::worksheet WsActivity_Struct;	// Activity_Struct 시트 객체
	xlnt::worksheet WsDebugInfo;		// dbuginfo 시트 객체
	
};

