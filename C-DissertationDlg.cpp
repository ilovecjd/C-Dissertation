
// C-DissertationDlg.cpp : 구현 파일
//
  
#include "stdafx.h"
#include "GlobalEnv.h"
#include "XLEzAutomation.h"
#include "C-Dissertation.h"
#include "C-DissertationDlg.h"
#include "Creator.h"
#include "Company.h"
#include "DlgProxy.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CCDissertationDlg 대화 상자


IMPLEMENT_DYNAMIC(CCDissertationDlg, CDialogEx);

CCDissertationDlg::CCDissertationDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_CDISSERTATION_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;

	//song 
	m_pGlobalEnv	= new GLOBAL_ENV;
	
}

CCDissertationDlg::~CCDissertationDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 NULL로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;

	//song
	if (m_pGlobalEnv != NULL)
		delete m_pGlobalEnv;
	//if (m_pActivityEnv != NULL)
		//delete m_pActivityEnv;
}

void CCDissertationDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CCDissertationDlg, CDialogEx)
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_CRETAT_PROJECT, &CCDissertationDlg::OnBnClickedCretatProject)
	ON_BN_CLICKED(IDC_SIMULATION_START, &CCDissertationDlg::OnBnClickedSimulationStart)
	ON_BN_CLICKED(IDC_LOAD, &CCDissertationDlg::OnBnClickedLoad)
	ON_BN_CLICKED(IDC_PRINT_EXCEL, &CCDissertationDlg::OnBnClickedPrintExcel)
	ON_BN_CLICKED(IDC_TEST, &CCDissertationDlg::OnBnClickedTest)
END_MESSAGE_MAP()


// CCDissertationDlg 메시지 처리기

BOOL CCDissertationDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 이 대화 상자의 아이콘을 설정합니다.  응용 프로그램의 주 창이 대화 상자가 아닐 경우에는
	//  프레임워크가 이 작업을 자동으로 수행합니다.
	SetIcon(m_hIcon, TRUE);			// 큰 아이콘을 설정합니다.
	SetIcon(m_hIcon, FALSE);		// 작은 아이콘을 설정합니다.

	// TODO: 여기에 추가 초기화 작업을 추가합니다.	
	srand((unsigned int)time(NULL));				// 난수 생성기 초기화

	m_pGlobalEnv->SimulationWeeks	= 4 * 36;		// 4주 x 36 개월
	m_pGlobalEnv->maxWeek		= 4 * 36 + 80;	//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	m_pGlobalEnv->WeeklyProb		= 1.25;
	m_pGlobalEnv->Hr_Init_H			= 2;
	m_pGlobalEnv->Hr_Init_M			= 1;
	m_pGlobalEnv->Hr_Init_L			= 3;
	m_pGlobalEnv->Hr_LeadTime		= 3;
	m_pGlobalEnv->Cash_Init			= 1000;
	m_pGlobalEnv->ProblemCnt		= 100;
	//m_pGlobalEnv->status			= 0;			// 프로그램의 동작 상태. 0:프로젝트 미생성, 1:프로젝트 생성,
	m_pGlobalEnv->ExpenseRate		= 1.6;
	return TRUE;  // 포커스를 컨트롤에 설정하지 않으면 TRUE를 반환합니다.
}

// 대화 상자에 최소화 단추를 추가할 경우 아이콘을 그리려면
//  아래 코드가 필요합니다.  문서/뷰 모델을 사용하는 MFC 응용 프로그램의 경우에는
//  프레임워크에서 이 작업을 자동으로 수행합니다.

void CCDissertationDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 그리기를 위한 디바이스 컨텍스트입니다.

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 클라이언트 사각형에서 아이콘을 가운데에 맞춥니다.
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 아이콘을 그립니다.
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// 사용자가 최소화된 창을 끄는 동안에 커서가 표시되도록 시스템에서
//  이 함수를 호출합니다.
HCURSOR CCDissertationDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 컨트롤러에서 해당 개체 중 하나를 계속 사용하고 있을 경우
//  사용자가 UI를 닫을 때 자동화 서버를 종료하면 안 됩니다.  이들
//  메시지 처리기는 프록시가 아직 사용 중인 경우 UI는 숨기지만,
//  UI가 표시되지 않아도 대화 상자는
//  남겨 둡니다.

void CCDissertationDlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CCDissertationDlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CCDissertationDlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CCDissertationDlg::CanExit()
{
	// 프록시 개체가 계속 남아 있으면 자동화 컨트롤러에서는
	//  이 응용 프로그램을 계속 사용합니다.  대화 상자는 남겨 두지만
	//  해당 UI는 숨깁니다.
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}



/* Excel R/W test Function
void CCDissertationDlg::OnBnClickedCretatProject()
{
	// Excel 자동화 객체 생성
	CXLEzAutomation xlAutomation;

	// Excel 파일 열기
	if (!xlAutomation.OpenExcelFile(_T("d:\\1.xlsx")))
	{
		MessageBox(_T("Failed to open Excel file."), _T("Error"), MB_OK | MB_ICONERROR);
		return;
	}

	// 데이터를 Excel에 쓰기 위해 배열 준비
	int posX = 4;
	int posY = 8;
	
	////////////////////////////////////////////////////////////////////////////
	int intArray2D[4][3] = { { 1, 2, 3 },{ 4, 5, 6 },{ 7, 8, 9 },{ 10, 11, 12 }, };
	int readIntArray2D[4][3] = { 0 };

	// 2D int 배열 데이터를 Excel에 쓰기
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, (int*)intArray2D, 4, 3);	
	// 2D int 배열 데이터를 Excel에서 읽기
	xlAutomation.ReadRangeToArray(PROJECT, posX, posY, (int*)readIntArray2D, 4, 3);
	{
		CString message;
		for (int r = 0; r < 4; r++) {
			for (int c = 0; c < 3; c++) {
				message.AppendFormat(_T("readIntArray2D[%d][%d] = %d\n"), r, c, readIntArray2D[r][c]);
			}
		}
		MessageBox(message, _T("Read 2D int array from Excel"), MB_OK);
	}
	
	////////////////////////////////////////////////////////////////////////////


	////////////////////////////////////////////////////////////////////////////
	CString readStrArray2D[3][1];
	CString strArray2D[3][1] = { _T("A1"), _T("A2"), _T("A3") };

	// 2D CString 배열 데이터를 Excel에 쓰기
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, (CString*)strArray2D, 3, 1);
	
	// 2D CString 배열 데이터를 Excel에서 읽기
	xlAutomation.ReadRangeToArray(PROJECT, posX, posY, (CString*)readStrArray2D, 3, 1);
	{
		CString message;
		for (int r = 0; r < 3; r++) {
			message.AppendFormat(_T("readStrArray2D[%d][0] = %s\n"), r, readStrArray2D[r][0]);
		}
		MessageBox(message, _T("Read 2D CString array from Excel"), MB_OK);
	}
	
	////////////////////////////////////////////////////////////////////////////


	////////////////////////////////////////////////////////////////////////////
	CString readStrArray2D2[1][3];
	CString strArray2D2[1][3] = { _T("2A1"), _T("2A2"), _T("2A3") };

	// 2D CString 배열 데이터를 Excel에 쓰기
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, (CString*)strArray2D2, 1, 3);

	// 2D CString 배열 데이터를 Excel에서 읽기
	xlAutomation.ReadRangeToArray(PROJECT, posX, posY, (CString*)readStrArray2D2, 1, 3);
	{
		CString message;
		for (int r = 0; r < 3; r++) {
			message.AppendFormat(_T("readStrArray2D2[0][%d] = %s\n"), r, readStrArray2D2[0][r]);
		}
		MessageBox(message, _T("Read 2D2 CString array from Excel"), MB_OK);
	}


	// 2D CString 배열 데이터를 Excel에 쓰기
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, (CString*)strArray2D2, 3, 1);

	// 2D CString 배열 데이터를 Excel에서 읽기
	xlAutomation.ReadRangeToArray(PROJECT, posX, posY, (CString*)readStrArray2D2, 3, 1);
	{
		CString message;
		for (int r = 0; r < 3; r++) {
			message.AppendFormat(_T("readStrArray2D2[0][%d] = %s\n"), r, readStrArray2D2[0][r]);
		}
		MessageBox(message, _T("Read 2D2 CString array from Excel"), MB_OK);
	}
	////////////////////////////////////////////////////////////////////////////



	////////////////////////////////////////////////////////////////////////////
	int readIntArray1D[3] = { 0 };
	int intArray1D[3] = { 10, 20, 30 };

	// 1D int 배열 데이터를 Excel에 쓰기
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, intArray1D, 3, 1);
	
	// 1D int 배열 데이터를 Excel에서 읽기
	xlAutomation.ReadRangeToArray(PROJECT, posX, posY, readIntArray1D, 3, 1);
	{
		CString message;
		for (int r = 0; r < 3; r++) {
			message.AppendFormat(_T("readIntArray1D[%d] = %d\n"), r, readIntArray1D[r]);
		}
		MessageBox(message, _T("Read 1D int array from Excel"), MB_OK);
	}



	// 1D int 배열 데이터를 Excel에 쓰기
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, intArray1D, 1, 3);

	// 1D int 배열 데이터를 Excel에서 읽기
	xlAutomation.ReadRangeToArray(PROJECT, posX, posY, readIntArray1D, 1, 3);
	{
		CString message;
		for (int r = 0; r < 3; r++) {
			message.AppendFormat(_T("readIntArray1D[%d] = %d\n"), r, readIntArray1D[r]);
		}
		MessageBox(message, _T("Read 1D int array from Excel"), MB_OK);
	}
	////////////////////////////////////////////////////////////////////////////
	

	////////////////////////////////////////////////////////////////////////////	
	CString readStrArray1D[3];
	CString strArray1D[3] = { _T("B1"), _T("B2"), _T("B3") };

	// 1D CString 배열 데이터를 Excel에 쓰기	
	xlAutomation.WriteArrayToRange(PROJECT, posX, posY, strArray1D, 3, 1);
	
	// 1D CString 배열 데이터를 Excel에서 읽기
	if (xlAutomation.ReadRangeToArray(PROJECT, posX, posY, readStrArray1D, 3, 1))
	{
		CString message;
		for (int r = 0; r < 3; r++) {
			message.AppendFormat(_T("readStrArray1D[%d] = %s\n"), r, readStrArray1D[r]);
		}
		MessageBox(message, _T("Read 1D CString array from Excel"), MB_OK);
	}
	else
	{
		MessageBox(_T("Failed to read 1D CString array from Excel."), _T("Error"), MB_OK | MB_ICONERROR);
	}
	////////////////////////////////////////////////////////////////////////////


	// Excel에서 다양한 데이터 유형 읽기
	int readIntValue;
	double readFloatValue; // Excel에서는 float이 double로 처리될 수 있으므로 double로 읽음
	double readDoubleValue;
	CString readStrValue;

	// 다양한 데이터 유형을 Excel에 쓰고 읽기
	xlAutomation.SetCellValue(PROJECT, posX, posY, 42);              // Integer
	if (xlAutomation.GetCellValue(PROJECT, posX, posY, &readIntValue)) {
		CString message;
		message.Format(_T("Read Integer value from Excel: %d"), readIntValue);
		MessageBox(message, _T("Read Integer"), MB_OK);
	}

	xlAutomation.SetCellValue(PROJECT, posX, posY, (float)3.14);     // Float (주의: Excel에서는 float이 double로 처리될 수 있음)
	if (xlAutomation.GetCellValue(PROJECT, posX, posY,  &readFloatValue)) {
		CString message;
		message.Format(_T("Read Float value from Excel: %lf"), readFloatValue);
		MessageBox(message, _T("Read Float"), MB_OK);
	}

	xlAutomation.SetCellValue(PROJECT, posX, posY, 2.71828);        // Double
	if (xlAutomation.GetCellValue(PROJECT, posX, posY, &readDoubleValue)) {
		CString message;
		message.Format(_T("Read Double value from Excel: %lf"), readDoubleValue);
		MessageBox(message, _T("Read Double"), MB_OK);
	}

	xlAutomation.SetCellValue(PROJECT, posX, posY, _T("문자"));     // CString
	if (xlAutomation.GetCellValue(PROJECT, posX, posY, &readStrValue)) {
		CString message;
		message.Format(_T("Read CString value from Excel: %s"), readStrValue);
		MessageBox(message, _T("Read CString"), MB_OK);
	}
}
*/


void CCDissertationDlg::OnBnClickedCretatProject()
{
	CCreator Creator;
	//PALL_ACT_TYPE pActType = new ALL_ACT_TYPE;
	int actTemp[] = { 50, 50, 2, 4, 2, 1, 60, 2, 40, 0, 0, 0, 0,
						20, 70, 5, 12, 2, 3, 50, 4, 50, 0, 0, 0, 0,
						20, 90, 13, 26, 2, 5, 50, 6, 50, 0, 0, 0, 0,
						8, 98, 27, 52, 2, 5, 40, 6, 60, 0, 0, 0, 0,
						2, 100, 53, 80, 2, 5, 30, 6, 70, 0, 0, 0, 0 };
	//memcpy(pActType, actTemp, 5*9);

	//PALL_ACTIVITY_PATTERN pActPattern = new ALL_ACTIVITY_PATTERN;
	int patternTemp[] = { 1, 100, 100, 0, 30, 70, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		1, 100, 100,80, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		2, 10, 20,	60, 40, 0, 0, 0, 10, 80, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		2, 20, 30,	80, 20, 0, 0, 0, 10, 70, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		3, 20, 30,	70, 30, 0, 0, 0, 10, 60, 30, 0, 0, 0, 60, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		4, 20, 30,	80, 20, 0, 0, 0, 10, 60, 30, 0, 0, 0, 50, 50, 0, 0, 0, 40, 60, 0, 0, 0, 0, 0 };
	//memcpy(pActPattern, patternTemp, 6*26);


	m_pGlobalEnv->SimulationWeeks = 4 * 36;		// 4주 x 36 개월
	m_pGlobalEnv->maxWeek = 4 * 36 + 80;	//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	m_pGlobalEnv->WeeklyProb = 1.25;
	m_pGlobalEnv->Hr_Init_H = 2;
	m_pGlobalEnv->Hr_Init_M = 2;
	m_pGlobalEnv->Hr_Init_L = 1;
	m_pGlobalEnv->Hr_LeadTime = 3;
	m_pGlobalEnv->Cash_Init = 3000;
	m_pGlobalEnv->ProblemCnt = 100;	
	m_pGlobalEnv->selectOrder = 1;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로
	m_pGlobalEnv->recruit = 20;		// 충원에 필요한 운영비 (몇주분량인가?)
	m_pGlobalEnv->layoff = 0;			// 감원에 필요한 운영비 (몇주분량인가?)
	
	
	m_pGlobalEnv->ExpenseRate = 1;
	//m_pGlobalEnv->profitRate = ;
				 
	m_pGlobalEnv->selectOrder = 0; //선택 순서  1: 먼저 발생한 순서대로 2 : 금액이 큰 순서대로 3 : 금액이 작은 순서대로
				 
	m_pGlobalEnv->recruit = 160;  // 작을수록 공격적인 인원 충원 144 : 시뮬레이션 끝까지 충원 없음
	m_pGlobalEnv->layoff = 0;  // 클수록 공격적인 인원 감축, 0 : 부도까지 인원 유지


	Creator.Init(m_pGlobalEnv, (ALL_ACT_TYPE*)&actTemp, (ALL_ACTIVITY_PATTERN*)&patternTemp);
	CString strFileName = L"d:\\test.anh";
	Creator.Save(strFileName);
	
}

void CCDissertationDlg::Decision(int id,BOOL shouldLoad, int result[3])
{
	//CCompany* company = new CCompany;
	//company->Init(m_pGlobalEnv, id, shouldLoad);

	//int j = 0;
	//while (j < m_pGlobalEnv->SimulationWeeks)
	//{
	//	if (FALSE == company->Decision(j))  // j번째 기간에 결정해야 할 일들		
	//		j = m_pGlobalEnv->SimulationWeeks + 1;

	//	j++;
	//}
	//
	//int profit = company->CalculateFinalResult();

	//if (company)
	//{
	//	delete company;
	//}
}

void CCDissertationDlg::OnBnClickedSimulationStart()
{
	for (int i = 0; i < 50; i++)
	{
		CCompany* company = new CCompany;
		CString strFileName = L"d:\\test.anh";
		company->Init(strFileName);

		company->m_GlobalEnv.Hr_Init_H = 3;//2
		company->m_GlobalEnv.Hr_Init_M = 3;//1
		company->m_GlobalEnv.Hr_Init_L = 3;//3
		
		company->m_GlobalEnv.selectOrder = 1; //선택 순서  1: 먼저 발생한 순서대로 2 : 금액이 큰 순서대로 3 : 금액이 작은 순서대로
		company->m_GlobalEnv.recruit = 24;  //0: 충원 없음. 작을수록 공격적인 인원 충원 144 : 시뮬레이션 끝까지 충원 없음
		company->m_GlobalEnv.layoff = 0;  // 0: 충원 없음 클수록 공격적인 인원 감축, 0 : 부도까지 인원 유지
		company->m_GlobalEnv.ExpenseRate = 1.2;// -j*0.1;

		company->ReInit();
		
		int j = 0;
		while (j < m_pGlobalEnv->SimulationWeeks)
		{
			if (FALSE == company->Decision(j))  // j번째 기간에 결정해야 할 일들		
				j = m_pGlobalEnv->SimulationWeeks + 1;

			j++;
		}

		int profit = company->CalculateFinalResult();
		CString fileName = _T("aa");
		//company->PrintResult(fileName);

		if (company) {
			delete company;
			company = NULL;
		}
	}
}


void CCDissertationDlg::OnBnClickedLoad()
{
	CCompany* company = new CCompany;
	CString strFileName = L"d:\\test.anh";
	company->Init(strFileName);
	company->PrintResult(strFileName);
	//company->PrintDBData();

	delete company;
	
}


void CCDissertationDlg::OnBnClickedPrintExcel()
{
	
	CXLEzAutomation* pXl;
	pXl = new CXLEzAutomation;

	pXl->OpenExcelFile(_T("d:\\1.xlsx"),_T("song"));
	
	int rows = 2;
	//int cols = m_GlobalEnv.maxWeek;

	CString strDBoardTitle[1][21] = {
		{ _T("주"), _T("누계"), _T("발주"),_T(""),_T("수입"),_T("지출"),_T(""),_T("투입"), _T("HR_H"), _T("HR_M"), _T("HR_L"),
		_T(""),_T("여유"), _T("HR_H"), _T("HR_M"), _T("HR_L"), _T(""),_T("총원"), _T("HR_H"), _T("HR_M"), _T("HR_L") }
	};
	pXl->WriteArrayToRange(WS_NUM_DEBUG_INFO, 2, 1, (CString*)strDBoardTitle, 18, 1); //세로로 출력
	pXl->SetRangeBorder(WS_NUM_DEBUG_INFO, 2, 1, 4, rows + 1, xlContinuous, xlThin, RGB(0, 0, 0));
	pXl->SetRangeBorder(WS_NUM_DEBUG_INFO, 7, 1, 9, rows + 1, xlContinuous, xlThin, RGB(0, 0, 0));
	pXl->SetRangeBorder(WS_NUM_DEBUG_INFO, 12, 1, 14, rows + 1, xlContinuous, xlThin, RGB(0, 0, 0));
	pXl->SetRangeBorder(WS_NUM_DEBUG_INFO, 17, 1, 19, rows + 1, xlContinuous, xlThin, RGB(0, 0, 0));

	delete pXl;
	/*pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 3, 2, m_orderTable[0], 1, cols);
	pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 4, 2, m_orderTable[1], 1, cols);

	int* pWeeks = new int[cols];
	for (int i = 0; i < cols; i++)
	{
		pWeeks[i] = i + 1;
	}
	pXl->WriteArrayToRange(WS_NUM_DASHBOARD, 2, 2, pWeeks, 1, cols);

	delete[] pWeeks;*/
}


void CCDissertationDlg::OnBnClickedTest()
{


	ALL_ACT_TYPE* actTemp = new ALL_ACT_TYPE;
	ALL_ACTIVITY_PATTERN* patternTemp = new ALL_ACTIVITY_PATTERN;
	DefaultParameters(actTemp, patternTemp);

	// 같은 프로젝트들로 여러가지 상황을 만들어 본다.
	// 주당 발생확률 1.25 지출이 1.2 일때 인원 변동
	int lastResunt[20 * 4 * 4 * 4][10] = { 0, };
	for (int i = 0; i < 20; i++)
	{
		m_pGlobalEnv->WeeklyProb = 1.25;// i * 0.1;// 1.25;
		CCreator Creator;
		Creator.Init(m_pGlobalEnv, actTemp, patternTemp);

		CString prarmFile;// = L"d:\\test00.anh";
		prarmFile.Format(_T("d:\\test_%0d.ahn"), i);
		Creator.Save(prarmFile);


		/*for (int j = 0; j < 50; j++)
		{*/
			



			for (int h = 0; h < 4; h++){
				for (int m = 0; m < 4; m++){
					for (int l = 0; l < 4; l++){
						int successCnt = 0;//성공횟수
						int successProfit = 0; //성공시 금액
						int failCnt = 0; // 실패 횟수
						int failMon = 0; // 실패 개월

						
						for (int j = 0; j < 50; j++) {
							if ((h + m + l) == 0) break;
							CCompany* company = new CCompany;
							company->Init(prarmFile);
							company->ReInit();
							company->m_GlobalEnv.Hr_Init_H = h;//2
							company->m_GlobalEnv.Hr_Init_M = m;//1
							company->m_GlobalEnv.Hr_Init_L = l;//3
							company->m_GlobalEnv.Cash_Init = (50 * h + 39 * m + 25 * l) * 4 * 6* 1.2; //인원수 대비 6개월
							company->m_GlobalEnv.selectOrder = 1; //선택 순서  1: 먼저 발생한 순서대로 2 : 금액이 큰 순서대로 3 : 금액이 작은 순서대로
							company->m_GlobalEnv.recruit = 0;  //0: 충원 없음. 작을수록 공격적인 인원 충원 144 : 시뮬레이션 끝까지 충원 없음
							company->m_GlobalEnv.layoff = 0;  // 0: 감원 없음 클수록 공격적인 인원 감축, 0 : 부도까지 인원 유지
							company->m_GlobalEnv.ExpenseRate = 1.2;// -j*0.1;

							int k = 0;
							while (k < m_pGlobalEnv->SimulationWeeks)
							{
								if (FALSE == company->Decision(k))  // j번째 기간에 결정해야 할 일들		
									k = m_pGlobalEnv->SimulationWeeks + 1;

								k++;
							}

							if (k > m_pGlobalEnv->SimulationWeeks)//실패
							{
								failCnt += 1;
								failMon += company->m_lastDecisionWeek;
							}
							else
							{
								successCnt +=1;//성공횟수 증가
								successProfit += company->CalculateFinalResult(); //성공시 금액
							}

							if (company) {
								delete company;
								company = NULL;
							}
						}/////
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][0] = i;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][1] = h;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][2] = m;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][3] = l;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][4] = successCnt;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][5] = successProfit;
						if(successCnt)
							lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][6] = successProfit / successCnt;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][7] = failCnt;
						lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][8] = failMon;
						if(failCnt)
							lastResunt[i * 4 * 4 * 4 + h * 4 * 4 + m * 4 + l][9] = failMon/failCnt;
					}
				}
			}
		}
	//}
	

	delete actTemp;
	delete patternTemp;
}

void CCDissertationDlg::DefaultParameters(ALL_ACT_TYPE* act, ALL_ACTIVITY_PATTERN* pattern)
{
	//PALL_ACT_TYPE pActType = new ALL_ACT_TYPE;
	int actTemp[] = { 50, 50, 2, 4, 2, 1, 60, 2, 40, 0, 0, 0, 0,
		20, 70, 5, 12, 2, 3, 50, 4, 50, 0, 0, 0, 0,
		20, 90, 13, 26, 2, 5, 50, 6, 50, 0, 0, 0, 0,
		8, 98, 27, 52, 2, 5, 40, 6, 60, 0, 0, 0, 0,
		2, 100, 53, 80, 2, 5, 30, 6, 70, 0, 0, 0, 0 };
	//memcpy(pActType, actTemp, 5*9);

	//PALL_ACTIVITY_PATTERN pActPattern = new ALL_ACTIVITY_PATTERN;
	int patternTemp[] = { 1, 100, 100, 0, 30, 70, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		1, 100, 100,80, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		2, 10, 20,	60, 40, 0, 0, 0, 10, 80, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		2, 20, 30,	80, 20, 0, 0, 0, 10, 70, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		3, 20, 30,	70, 30, 0, 0, 0, 10, 60, 30, 0, 0, 0, 60, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
		4, 20, 30,	80, 20, 0, 0, 0, 10, 60, 30, 0, 0, 0, 50, 50, 0, 0, 0, 40, 60, 0, 0, 0, 0, 0 };
	//memcpy(pActPattern, patternTemp, 6*26);

	*act = *((ALL_ACT_TYPE*)actTemp);
	*pattern = *((ALL_ACTIVITY_PATTERN*)patternTemp);

	m_pGlobalEnv->SimulationWeeks = 4 * 36;		// 4주 x 36 개월
	m_pGlobalEnv->maxWeek = 4 * 36 + 80;	//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	m_pGlobalEnv->WeeklyProb = 1.25;
	m_pGlobalEnv->Hr_Init_H = 2;
	m_pGlobalEnv->Hr_Init_M = 2;
	m_pGlobalEnv->Hr_Init_L = 1;
	m_pGlobalEnv->Hr_LeadTime = 3;
	m_pGlobalEnv->Cash_Init = 3000;
	m_pGlobalEnv->ProblemCnt = 100;
	m_pGlobalEnv->selectOrder = 1;	// 선택 순서  1: 먼저 발생한 순서대로 2: 금액이 큰 순서대로 3: 금액이 작은 순서대로
	m_pGlobalEnv->recruit = 20;		// 충원에 필요한 운영비 (몇주분량인가?)
	m_pGlobalEnv->layoff = 0;			// 감원에 필요한 운영비 (몇주분량인가?)

	m_pGlobalEnv->ExpenseRate = 1.2;
	m_pGlobalEnv->selectOrder = 0; //선택 순서  1: 먼저 발생한 순서대로 2 : 금액이 큰 순서대로 3 : 금액이 작은 순서대로

	m_pGlobalEnv->recruit = 160;  // 작을수록 공격적인 인원 충원 144 : 시뮬레이션 끝까지 충원 없음
	m_pGlobalEnv->layoff = 0;  // 클수록 공격적인 인원 감축, 0 : 부도까지 인원 유지
}
