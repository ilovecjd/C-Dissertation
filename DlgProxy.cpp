
// DlgProxy.cpp : 구현 파일
//

#include "stdafx.h"
#include "C-Dissertation.h"
#include "DlgProxy.h"
#include "C-DissertationDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CCDissertationDlgAutoProxy

IMPLEMENT_DYNCREATE(CCDissertationDlgAutoProxy, CCmdTarget)

CCDissertationDlgAutoProxy::CCDissertationDlgAutoProxy()
{
	EnableAutomation();
	
	// 자동화 개체가 활성화되어 있는 동안 계속 응용 프로그램을 실행하기 위해 
	//	생성자에서 AfxOleLockApp를 호출합니다.
	AfxOleLockApp();

	// 응용 프로그램의 주 창 포인터를 통해 대화 상자에 대한
	//  액세스를 가져옵니다.  프록시의 내부 포인터를 설정하여
	//  대화 상자를 가리키고 대화 상자의 후방 포인터를 이 프록시로
	//  설정합니다.
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CCDissertationDlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CCDissertationDlg)))
		{
			m_pDialog = reinterpret_cast<CCDissertationDlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CCDissertationDlgAutoProxy::~CCDissertationDlgAutoProxy()
{
	// 모든 개체가 OLE 자동화로 만들어졌을 때 응용 프로그램을 종료하기 위해
	// 	소멸자가 AfxOleUnlockApp를 호출합니다.
	//  이러한 호출로 주 대화 상자가 삭제될 수 있습니다.
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CCDissertationDlgAutoProxy::OnFinalRelease()
{
	// 자동화 개체에 대한 마지막 참조가 해제되면
	// OnFinalRelease가 호출됩니다.  기본 클래스에서 자동으로 개체를 삭제합니다.
	// 기본 클래스를 호출하기 전에 개체에 필요한 추가 정리 작업을
	// 추가하십시오.

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CCDissertationDlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CCDissertationDlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 참고: IID_ICDissertation에 대한 지원을 추가하여
//  VBA에서 형식 안전 바인딩을 지원합니다.
//  이 IID는 .IDL 파일에 있는 dispinterface의 GUID와 일치해야 합니다.

// {2D2AA0A4-6E81-4BBB-AB37-F0F477771B5E}
static const IID IID_ICDissertation =
{ 0x2D2AA0A4, 0x6E81, 0x4BBB, { 0xAB, 0x37, 0xF0, 0xF4, 0x77, 0x77, 0x1B, 0x5E } };

BEGIN_INTERFACE_MAP(CCDissertationDlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CCDissertationDlgAutoProxy, IID_ICDissertation, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 매크로가 이 프로젝트의 StdAfx.h에 정의됩니다.
// {DD252F81-78B3-44B6-869E-F6496C136DA1}
IMPLEMENT_OLECREATE2(CCDissertationDlgAutoProxy, "CDissertation.Application", 0xdd252f81, 0x78b3, 0x44b6, 0x86, 0x9e, 0xf6, 0x49, 0x6c, 0x13, 0x6d, 0xa1)


// CCDissertationDlgAutoProxy 메시지 처리기
