
// C-DissertationDlg.h : 헤더 파일
//

#pragma once

class CCDissertationDlgAutoProxy;


// CCDissertationDlg 대화 상자
class CCDissertationDlg : public CDialogEx
{
	DECLARE_DYNAMIC(CCDissertationDlg);
	friend class CCDissertationDlgAutoProxy;

// 생성입니다.
public:
	CCDissertationDlg(CWnd* pParent = NULL);	// 표준 생성자입니다.
	virtual ~CCDissertationDlg();

// 대화 상자 데이터입니다.
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_CDISSERTATION_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 지원입니다.


// 구현입니다.
protected:
	CCDissertationDlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 생성된 메시지 맵 함수
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedCretatProject();
};
