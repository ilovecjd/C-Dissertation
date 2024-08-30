
// C-DissertationDlg.cpp : 구현 파일
//
  
#include "stdafx.h"
#include "XLEzAutomation.h"
#include "C-Dissertation.h"
#include "C-DissertationDlg.h"
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
	m_pGlobalEnv->SimulationWeeks	= 4 * 36;		// 4주 x 36 개월
	m_pGlobalEnv->Hr_TableSize		= 4 * 36 + 80;	//  maxTableSize 최대 80주(18개월)간 진행되는 프로젝트를 시뮬레이션 마지막에 기록할 수도 있다.
	m_pGlobalEnv->WeeklyProb		= 1.25;
	m_pGlobalEnv->Hr_Init_H			= 13;
	m_pGlobalEnv->Hr_Init_M			= 21;
	m_pGlobalEnv->Hr_Init_L			= 6;
	m_pGlobalEnv->Hr_LeadTime		= 3;
	m_pGlobalEnv->Cash_Init			= 1000;
	m_pGlobalEnv->ProblemCnt		= 100;
	m_pGlobalEnv->status			= 0;			// 프로그램의 동작 상태. 0:프로젝트 미생성, 1:프로젝트 생성,

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
	CCompany* company = new CCompany; 
	company->Init(m_pGlobalEnv, 1, TRUE);
}

void CCDissertationDlg::OnBnClickedSimulationStart()
{
	// TODO: Add your control notification handler code here
}
