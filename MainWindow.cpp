#include "stdafx.h"

#include "../Include//vba.h"
#include "../Include/Vbaext.h"
#include "../Include/msotl.h"

#include "VBAWIN_I.H"
#include "VBAWIN.H"
#include "Resource.h"

#include "../Include/language_token.h"

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

HRESULT STDMETHODCALLTYPE VBAWINApp::Alert( 
	/* [in] */ BSTR message)
{
	MessageBoxW(pmainwin->m_hWnd, message?message:L"", L"VBAWINAPP", MB_OK);
	return S_OK;
}

const IID IID_VBEOBJ = {
	0x0002e166, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 }
};

HRESULT STDMETHODCALLTYPE VBAWINApp::get_VBE( 
	/* [retval][out] */ /* external definition not present */ VBE **ppVbe) 
{
	static const wchar_t *const VBE_INIT_FAILE = L"You must show Visual Basic Editor at least once!";
	__asm cmp ppVbe, 0
	__asm mov eax, 80070057h // E_INVALIDARG
	__asm je label_RET
	// m_hVbe
	__asm mov ecx, this
	__asm mov eax, m_hVbe
	__asm add eax, ecx	
	// global VBE Object
	__asm mov eax, [eax]
	__asm add eax, 24CA3Ch
	// vftable
	__asm mov eax, [eax]
	// the vbe must be show at least once
	__asm test eax, eax
	__asm je label_RET_E_FAIL	
	// ppv
	__asm push ppVbe
	// riid
	__asm push OFFSET IID_VBEOBJ
	// this
	__asm push eax
	// QueryInterface
	__asm mov ecx, [eax]
	__asm call [ecx]
	__asm jmp label_RET
label_RET_E_FAIL:
	__asm push VBE_INIT_FAILE
	__asm push this
	__asm call Alert
	__asm mov eax, 80004005h // E_FAIL
label_RET: ;
}


VBAWINMAINWIN::VBAWINMAINWIN(void)
{
	mp_Vba = NULL;
	m_cMsgLoop = 0;
	m_cFirstMsgLoop = 0;
	m_fHostIsActiveComponent = FALSE;
	m_AppObject.AddRef();
	mp_HostAddins = NULL;
}

VBAWINMAINWIN::~VBAWINMAINWIN(void)
{
	mp_Vba = NULL;
}

int VBAWINMAINWIN::RunMessageLoop(void)
{
	BOOL fAborted;
	return MessageLoop((VBALOOPREASON)0, &fAborted, TRUE);
}

int VBAWINMAINWIN::MessageLoop(VBALOOPREASON loopreason, BOOL* pfAborted, BOOL fPushedByHost)
{
	int nExitCode = 0;

	MSG msg;
	BOOL fContinue, fConsumed;
	int idlecount;
	msg.message = 0;
	m_cMsgLoop++;

	if (fPushedByHost && 0 == m_cFirstMsgLoop)
	{
		m_cFirstMsgLoop = m_cMsgLoop;
	}

	CComPtr<IVbaCompManager> compmgr;
	while (TRUE)
	{
		if(!compmgr){
			mp_Vba->GetCompManager(&compmgr);
		}

		if (compmgr)
		{
			fContinue = TRUE;
			compmgr->ContinueMessageLoop(fPushedByHost, &fContinue);
			if (!fContinue)
			{
				if (pfAborted)
					*pfAborted = TRUE;
				break;
			}
		}

		if (!PeekMessage(&msg, NULL, NULL, NULL, PM_REMOVE))
		{				
			// Although OnIdle request more idle time,
			// we cann't meet it all the time.
			idlecount =0;
			while(OnIdle() && idlecount < 16) 
				idlecount++;

			if (!PeekMessage(&msg, NULL, NULL, NULL, PM_NOREMOVE))
			{
				if(compmgr)
					compmgr->OnWaitForMessage();
				WaitMessage();				
			}
			continue;
		}

		if (msg.message == WM_QUIT)
		{
			if (m_cMsgLoop > 1)
				PostQuitMessage(msg.wParam);
			else
				nExitCode = msg.wParam;
			break;
		}

		fConsumed = FALSE;

		if (compmgr)
			compmgr->PreTranslateMessage(&msg, &fConsumed);
		
		if (fConsumed)
			continue;
		
		if (!m_fHostIsActiveComponent)
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
			continue;
		}
		
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	if (fPushedByHost && m_cMsgLoop == m_cFirstMsgLoop)
		m_cFirstMsgLoop = 0;

	m_cMsgLoop--;

	//if (m_cMsgLoop == m_cFirstMsgLoop)
		//ApcHost.StackUnwound();

	return nExitCode;
}

BOOL VBAWINMAINWIN::OnIdle(BOOL fPeriodic, BOOL fPriority, BOOL fNonPeriodic)
{
	BOOL pfContinue = FALSE;

	CComPtr<IVbaCompManager> compmgr;
	mp_Vba->GetCompManager(&compmgr);
	if (!compmgr) return FALSE;

	// if we have periodic idle tasks that to be done, 
	// do them and return TRUE.
	if (fPeriodic) 
	{
		compmgr->DoIdle(VBAIDLEFLAG_Periodic, &pfContinue);
		if (pfContinue) return TRUE;
	}

	// if any high priority idle tasks need to be done,
	// do it, and return TRUE.
	if (fPriority)
	{
		compmgr->DoIdle(VBAIDLEFLAG_Priority, &pfContinue);
		if (pfContinue) return TRUE;
	}

	// if any lower priority idle tasks need to be done,
	// do it, and return TRUE.
	if (fNonPeriodic)
	{
		compmgr->DoIdle(VBAIDLEFLAG_NonPeriodic, &pfContinue);
		if (pfContinue) return TRUE;
	}
	return FALSE;
}

LRESULT VBAWINMAINWIN::OnClose(UINT umsg, WPARAM wp, LPARAM lp, BOOL&)
{
	BOOL fCanTerm;
	mp_Vba->QueryTerminate(&fCanTerm);
	if(fCanTerm)
	{
		DestroyWindow();
		return 0;
	}
	return TRUE;
}

LRESULT VBAWINMAINWIN::OnDestroy(UINT umsg, WPARAM wp, LPARAM lp, BOOL&)
{
	PostQuitMessage(0);	
	return 0;
}

LRESULT VBAWINMAINWIN::OnActivate(UINT umsg, WPARAM wp, LPARAM lp, BOOL&)
{
	CComVbaCompMgrPtr ptr;
	if(LOWORD(wp) == WA_INACTIVE)
		m_fHostIsActiveComponent = FALSE;
	else
	{
		mp_Vba->GetCompManager(&ptr);
		if(ptr)
			ptr->OnHostActivate();
		m_fHostIsActiveComponent = TRUE;
	}
	return 0;
}

LRESULT VBAWINMAINWIN::OnEnable(UINT umsg, WPARAM wp, LPARAM lp, BOOL&)
{
	CComVbaCompMgrPtr ptr;
	if(!m_fWmEnableFromCompMgr)
	{
		mp_Vba->GetCompManager(&ptr);
		if(ptr)
			ptr->OnHostEnterState(VBACOMPSTATE_Modal, wp == 0);
	}
	return 0;
}

extern "C" const IID SID_SVbaMainUI = {
	0x759d0501, 0xd979, 0x11ce, { 0x84,0xec,0x00,0xaa,0x00,0x61,0x4f,0x3e }
};
const IID IID_CommandBarButton = {
	0x000C030EL,0x0000,0x0000, { 0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46 }
};
const IID IID_CommandBarControl = {
	0x000C0308L,0x0000,0x0000, { 0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46 }
};

LRESULT VBAWINMAINWIN::OnCommand(UINT umsg, WPARAM wp, LPARAM lp, BOOL&)
{
	if(HIWORD(wp) == 0)
	{
		switch(LOWORD(wp))
		{
		case ID_CMD_VBE:{
			CComQIPtr<IServiceProvider> srvprov = mp_Vba;
			CComPtr<IUIElement> vbaui;			
			srvprov->QueryService(SID_SVbaMainUI, IID_IUIElement, (void**)&vbaui);
			vbaui->Show();
			} break;
		case ID_CMD_ADDINS:
			mp_HostAddins->ShowAddInsDialog(m_hWnd);
			break;
		__case(ID_CMD_MACROS)
			CComPtr<VBE> __v;
			HRESULT hr = m_AppObject.get_VBE(&__v);
			if(SUCCEEDED(hr)) {
				CComPtr<CommandBars> __b;
				hr = __v->get_CommandBars(&__b);
				if(SUCCEEDED(hr)) {
					int cnt;
					__b->get_Count(&cnt);
					VARIANT var;
					WCHAR buf[4];
					var.vt = VT_BSTR;
					CComPtr<CommandBar> __c;
					var.bstrVal = SysAllocString(L"Tools");
					hr = __b->get_Item(var, &__c);
					SysFreeString(var.bstrVal);
					if(SUCCEEDED(hr)) {
						long cldcnt;
						hr = __c->get_accChildCount(&cldcnt);
						var.vt= VT_I4;
						CComPtr<IDispatch> _cc1;
						var.uintVal = 3;
						hr = __c->get_accChild(var, &_cc1);
						if(SUCCEEDED(hr)) {
							CComPtr<CommandBarControl> btn;
							hr = _cc1->QueryInterface(IID_CommandBarControl, (void**)&btn);
							if(SUCCEEDED(hr)) {
								var.uintVal = CHILDID_SELF;
								btn->accDoDefaultAction(var);
							}
							hr = hr;
						}
					}
				}
			}
			__cbreak;
		}
	}
	return 0;
}
