﻿
// C-DissertationDlg.cpp : 구현 파일
//

#include "stdafx.h"
#include "XLEzAutomation.h"
#include "C-Dissertation.h"
#include "C-DissertationDlg.h"
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
}

CCDissertationDlg::~CCDissertationDlg()
{
	// 이 대화 상자에 대한 자동화 프록시가 있을 경우 이 대화 상자에 대한
	//  후방 포인터를 NULL로 설정하여
	//  대화 상자가 삭제되었음을 알 수 있게 합니다.
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
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


/*
void CCDissertationDlg::OnBnClickedCretatProject()
{
	// 엑셀 자동화 객체 생성
	CXLAutomation* Xl = new CXLAutomation;

	// 엑셀 파일 열기
	if (!Xl->OpenExcelFile(_T("d:\\1.xlsx")))
	{
		MessageBox(_T("엑셀 파일을 열 수 없습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		delete Xl;
		return;
	}

	// 데이터를 저장할 2차원 배열 선언
	const int rows = 3;
	const int cols = 5;
	int dataArray[rows][cols] = { 0 };

	// 엑셀 시트의 특정 범위를 읽어서 배열에 저장
	int startRow = 8, startCol = 6, endRow = 8+2, endCol = 6+4;
	if (!Xl->ReadRangeToArray(PROJECT, startRow, startCol, endRow, endCol, (int*)dataArray, rows, cols))
	{
		MessageBox(_T("엑셀 데이터 범위를 읽어올 수 없습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		Xl->ReleaseExcel();
		delete Xl;
		return;
	}

	// 가져온 데이터 배열 확인 (예시로 출력)
	for (int i = 0; i < rows; i++)
	{
		for (int j = 0; j < cols; j++)
		{
			CString str;
			str.Format(_T("dataArray[%d][%d] = %d"), i, j, dataArray[i][j]);
			MessageBox(str, _T("Data"), MB_OK);
		}
	}


	// 엑셀 리소스 해제
	Xl->ReleaseExcel();
	delete Xl;
}*/


/*
void CCDissertationDlg::OnBnClickedCretatProject()
{
	// 엑셀 자동화 객체 생성
	CXLAutomation* Xl = new CXLAutomation;

	// 엑셀 파일 열기
	if (!Xl->OpenExcelFile(_T("d:\\1.xlsx")))
	{
		MessageBox(_T("엑셀 파일을 열 수 없습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		delete Xl;
		return;
	}

	int intValue = 0;
	double doubleValue = 0.0;
	CString strValue;

	// 정수 값을 테스트하기 위해 셀 값을 가져오기
	if (Xl->GetCellValueInt(PROJECT, 1, 1, &intValue))
	{
		CString message;
		message.Format(_T("Integer value in cell (1,1) is: %d"), intValue);
		MessageBox(message, _T("Integer Value"), MB_OK);
	}
	else
	{
		MessageBox(_T("셀에서 정수 값을 가져오지 못했습니다."), _T("Error"), MB_OK | MB_ICONERROR);
	}

	// 실수 값을 테스트하기 위해 셀 값을 가져오기
	if (Xl->GetCellValueDouble(PROJECT, 2, 2, &doubleValue))
	{
		CString message;
		message.Format(_T("Double value in cell (2,2) is: %f"), doubleValue);
		MessageBox(message, _T("Double Value"), MB_OK);
	}
	else
	{
		MessageBox(_T("셀에서 실수 값을 가져오지 못했습니다."), _T("Error"), MB_OK | MB_ICONERROR);
	}

	// 문자열 값을 테스트하기 위해 셀 값을 가져오기
	if (Xl->GetCellValueCString(PROJECT, 3, 3, &strValue))
	{
		CString message;
		message.Format(_T("String value in cell (3,3) is: %s"), strValue);
		MessageBox(message, _T("String Value"), MB_OK);
	}
	else
	{
		MessageBox(_T("셀에서 문자열 값을 가져오지 못했습니다."), _T("Error"), MB_OK | MB_ICONERROR);
	}

	// 엑셀 리소스 해제
	Xl->ReleaseExcel();
	delete Xl;
}
*/

void CCDissertationDlg::OnBnClickedCretatProject()
{
	// CXLEzAutomation 객체 생성
	CXLEzAutomation* xlAutomation = new CXLEzAutomation(TRUE); // Excel을 보이도록 생성

															   // 엑셀 파일 열기
	if (!xlAutomation->OpenExcelFile(_T("d:\\1.xlsx")))
	{
		MessageBox(_T("엑셀 파일을 열 수 없습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		delete xlAutomation;
		return;
	}

	// 다양한 유형의 데이터를 Excel 셀에 설정하고 읽기
	try
	{
		// 정수 값 설정 및 읽기
		int intValue = 42;
		xlAutomation->SetCellValue(PROJECT, 1, 1, intValue); // 1행 1열 (A1 셀)에 42 설정
		int readIntValue;
		if (xlAutomation->GetCellValue(PROJECT, 1, 1, &readIntValue))
		{
			CString msg;
			msg.Format(_T("Read integer value: %d"), readIntValue);
			MessageBox(msg, _T("Info"), MB_OK);
		}
		else
		{
			MessageBox(_T("정수 값을 읽어오는데 실패했습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		}

		// 문자열 값 설정 및 읽기
		CString strValue = _T("Hello Excel");
		xlAutomation->SetCellValue(PROJECT, 2, 1, strValue); // 1행 2열 (B1 셀)에 "Hello Excel" 설정
		CString readStrValue;
		if (xlAutomation->GetCellValue(PROJECT, 2, 1, &readStrValue))
		{
			CString msg;
			msg.Format(_T("Read string value: %s"), readStrValue);
			MessageBox(msg, _T("Info"), MB_OK);
		}
		else
		{
			MessageBox(_T("문자열 값을 읽어오는데 실패했습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		}

		// 실수 값 설정 및 읽기
		double dblValue = 3.14159;
		xlAutomation->SetCellValue(PROJECT, 3, 1, dblValue); // 1행 3열 (C1 셀)에 3.14159 설정
		double readDblValue;
		if (xlAutomation->GetCellValue(PROJECT, 3, 1, &readDblValue))
		{
			CString msg;
			msg.Format(_T("Read double value: %f"), readDblValue);
			MessageBox(msg, _T("Info"), MB_OK);
		}
		else
		{
			MessageBox(_T("실수 값을 읽어오는데 실패했습니다."), _T("Error"), MB_OK | MB_ICONERROR);
		}
	}
	catch (const std::exception& e)
	{
		MessageBox(CString(e.what()), _T("Error"), MB_OK | MB_ICONERROR);
	}

	// 엑셀 리소스 해제
	xlAutomation->ReleaseExcel();
	delete xlAutomation;
}




