#ifndef ___VBAWIN_H_INCLUDE___
#define ___VBAWIN_H_INCLUDE___

extern "C" const IID IID_IVbaSite;
extern "C" const IID IID_IVbaCompManagerSite;
extern "C" const IID IID_IVbaProjectSite;

typedef CComPtr<IVba> CComVbaPtr;
typedef CComPtr<IVbaCompManager> CComVbaCompMgrPtr;

extern class VBAWINMAINWIN *pmainwin;

class VBAWINApp : 
	public CComObjectRoot,
	public IDispatchImpl<IVBAWINApp, &IID_IVBAWINApp>
{
public:
	//DECLARE_CLASSFACTORY()

	//DECLARE_REGISTRY_RESOURCEID(IDR_VBAWIN)
	DECLARE_NO_REGISTRY()

	DECLARE_NOT_AGGREGATABLE(VBAWINApp)

	BEGIN_COM_MAP(VBAWINApp)
		COM_INTERFACE_ENTRY_IID(IID_IVBAWINApp, IVBAWINApp)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP();

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
		/* [retval][out] */ IVBAWINApp **ppApp)
	{
		return QueryInterface(IID_IVBAWINApp, (void**)ppApp);
	}

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 		
		/* [retval][out] */ IVBAWINApp **ppApp)
	{
		*ppApp = NULL;
		return S_FALSE;
	}

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Name( 
		/* [retval][out] */ LPBSTR name)
	{
		*name = SysAllocString(L"SAYUMI SOUL");
		return S_OK;
	}

	virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE Alert( 
		/* [in] */ BSTR message);

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_VBE( 
		/* [retval][out] */ /* external definition not present */ VBE **ppVbe);

	VBAWINApp()
#if VBE7
	//: m_hVbe(_T("VBE7.DLL"))
#else
	//: m_hVbe(_T("VBE6.DLL"))
#endif
	{
		m_hVbe = GetModuleHandleW(L"VBE7.DLL");
	}

#if VBE7
	//MODULEHANDLE m_hVbe;
	HANDLE m_hVbe;
#else
	//MODULEHANDLE m_hVbe;
	HANDLE m_hVbe;
#endif
};

class VBAWINMAINWIN :
	public CWindowImpl<VBAWINMAINWIN, CWindow, CFrameWinTraits>
{
	BEGIN_MSG_MAP(VBAWINMAINWIN)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_ACTIVATE, OnActivate)
		MESSAGE_HANDLER(WM_ENABLE, OnEnable)
		MESSAGE_HANDLER(WM_COMMAND, OnCommand)
	END_MSG_MAP()

public:
	VBAWINMAINWIN(void);
	~VBAWINMAINWIN(void);

	int RunMessageLoop(void);
	int MessageLoop(VBALOOPREASON loopreason, BOOL* pfAborted, BOOL fPushedByHost);
	BOOL/*Continue*/ OnIdle(BOOL fPeriodic = TRUE, BOOL fPriority = TRUE, BOOL fNonPeriodic = TRUE);
	BOOL m_fHostIsActiveComponent;
	BOOL m_fWmEnableFromCompMgr;

	IVba *mp_Vba;
	CComObject<VBAWINApp> m_AppObject;
	IVbaHostAddins* mp_HostAddins;

private:
	LRESULT OnClose(UINT umsg, WPARAM wp, LPARAM lp, BOOL&);
	LRESULT OnDestroy(UINT umsg, WPARAM wp, LPARAM lp, BOOL&);
	LRESULT OnActivate(UINT umsg, WPARAM wp, LPARAM lp, BOOL&);
	LRESULT OnEnable(UINT umsg, WPARAM wp, LPARAM lp, BOOL&);
	LRESULT OnCommand(UINT umsg, WPARAM wp, LPARAM lp, BOOL&);

	//===================
	int m_cMsgLoop;
	int m_cFirstMsgLoop;
};

class VBACOMPMGRSITE :
	public CComObjectRoot,
	public IVbaCompManagerSite
{
	BEGIN_COM_MAP(VBACOMPMGRSITE)
		COM_INTERFACE_ENTRY_IID(IID_IVbaCompManagerSite, IVbaCompManagerSite)
	END_COM_MAP()

public:
	VBACOMPMGRSITE(void) : modalcount(0) { }

	// *** IVbaCompManagerSite methods ***
	STDMETHOD(Notify)(VBACMSITENOTIFY vbacmsn) {
		if(vbacmsn == VBACMSN_LostActivation)
			pmainwin->m_fHostIsActiveComponent = FALSE;
		return S_OK;
	}

	int modalcount;
	STDMETHOD(EnterState)(THIS_ VBACOMPSTATE vbacompstate, BOOL fEnter) {
		if(vbacompstate == VBACOMPSTATE_Modal)
		{
			if(fEnter)
			{
				modalcount++;
				if(modalcount == 1) {
					pmainwin->m_fWmEnableFromCompMgr = TRUE;
					pmainwin->EnableWindow(FALSE);
					pmainwin->m_fWmEnableFromCompMgr = FALSE;
				}
			}
			else if(modalcount)
			{
				modalcount--;
				if(modalcount == 0) {
					pmainwin->m_fWmEnableFromCompMgr = TRUE;
					pmainwin->EnableWindow(TRUE);
					pmainwin->m_fWmEnableFromCompMgr = FALSE;
				}
			}
		}
		return S_OK;
	}

	STDMETHOD(PushMessageLoop)(THIS_ VBALOOPREASON vbaloopreason, BOOL *pfAborted) {
		pmainwin->MessageLoop(vbaloopreason, pfAborted, FALSE);
		return S_OK;
	}
	STDMETHOD(ContinueIdle)(THIS_ BOOL *pfContinue) {
		return E_NOTIMPL;
	}
	STDMETHOD(DoIdle)(THIS) {
		pmainwin->OnIdle();
		return S_OK;
	}
};

void DoDetour(
	HMODULE hVbeDll = NULL
);

//{000C060B-0000-0000-C000-000000000046}
extern "C" const IID IID_IMsoComponentManager;
//{000C0601-0000-0000-C000-000000000046}
extern "C" const IID SID_SMsoComponentManager;
//{000C0600-0000-0000-C000-000000000046}
extern "C" const IID IID_IMsoComponent;

typedef enum {
	msocstateModal = 1,
	msocstateRedrawOff = 2,
	msocstateWarningsOff = 3,
	msocstateRecording = 4
} Msocstate;

// Specifies registration information for a component.
typedef struct _MSOCRINFO {
	ULONG cbSize;
	ULONG uIdleTimeInterval;
	DWORD grfcrf;
	DWORD grfcadvf;
} MSOCRINFO;

// Specifies information for a component host.
typedef struct _MSOCHOSTINFO {
	ULONG cbSize;
	DWORD grfchostf;
} MSOCHOSTINFO;

#define msoidlefPeriodic     1
#define msoidlefNonPeriodic  2
#define msoidlefPriority     4
#define msoidlefAll          0xFFFFFFFF
#define MSOIDLEFLAGS DWORD

typedef enum {
	msoloopFocusWait = 1,
	msoloopDoEvents = 2,
	msoloopDebug = 3,
	msoloopModalForm = 4,
	msoloopModalAlert = 5
} Msoloop;

#define msocWindowFrameToplevel  0
#define msocWindowFrameOwner     1
#define msocWindowComponent      2
#define msocWindowDlgOwner       3
#define MSOCWINDOWTYPE DWORD

struct __declspec(uuid("000C0600-0000-0000-C000-000000000046"))
IMsoComponent : public IUnknown
{
	// always return s TRUE
	virtual BOOL STDMETHODCALLTYPE FDebugMessage(
		HANDLE hinst, UINT message, WPARAM wParam, LPARAM lParam ) = 0;
	/* Enables components to process a message before the message
	 * is translated and is dispatched. Components can use the 
	 * TranslateAccelerator function, use the IsDialogMessage function, 
	 * modify the pMsg parameter, or take some other action. */
	virtual BOOL STDMETHODCALLTYPE FPreTranslateMessage(MSG *pMsg) = 0;
	virtual void STDMETHODCALLTYPE OnEnterState(Msocstate uStateID, BOOL fEnter) = 0;
	virtual void STDMETHODCALLTYPE OnAppActivate(BOOL fActive, DWORD dwOtherThreadID) = 0;
	virtual void STDMETHODCALLTYPE OnLoseActivation(void) = 0;
	virtual void STDMETHODCALLTYPE OnActivationChange(IMsoComponent *pic, 
		BOOL fSameComponent, const MSOCRINFO *pcrinfo, BOOL fHostIsActivating, 
		const MSOCHOSTINFO *pchostinfo, DWORD dwReserved) = 0;
	virtual BOOL STDMETHODCALLTYPE FDoIdle(MSOIDLEFLAGS grfidlef) = 0;
	virtual BOOL STDMETHODCALLTYPE FContinueMessageLoop(
		Msoloop uReason,void *pvLoopData, MSG *pMsgPeeked) = 0;
	virtual BOOL STDMETHODCALLTYPE FQueryTerminate(BOOL fPromptUser) = 0;
	virtual void STDMETHODCALLTYPE Terminate(void) = 0;
	virtual HWND STDMETHODCALLTYPE HwndGetWindow(MSOCWINDOWTYPE dwWhich,DWORD dwReserved) = 0;
};

#define msoccontextAll     0
#define msoccontextMine    1
#define msoccontextOthers  2
#define MSOCCONTEXT DWORD

#define msogacActive             0
#define msogacTracking           1
#define msogacTrackingOrActive   2
#define msogacModal              3
#define MSOGAC DWORD

struct __declspec(uuid("000C060B-0000-0000-C000-000000000046"))
IMsoComponentManager : public IServiceProvider
{
	// always returns TRUE
	virtual BOOL STDMETHODCALLTYPE FDebugMessage(
		ULONG dwReserved, ULONG message, DWORD_PTR wParam, DWORD_PTR lParam) = 0;
	virtual BOOL STDMETHODCALLTYPE FRegisterComponent(
		IMsoComponent *piComponent, const MSOCRINFO *pcrinfo, DWORD_PTR *pdwComponentID) = 0;
	virtual HRESULT STDMETHODCALLTYPE  FRevokeComponent(DWORD_PTR dwComponentID) = 0;
	virtual HRESULT STDMETHODCALLTYPE FUpdateComponentRegistration(
		DWORD_PTR dwComponentID, const MSOCRINFO *pcrinfo) = 0;
	virtual HRESULT STDMETHODCALLTYPE FOnComponentActivate(DWORD_PTR dwComponentID) = 0;
	virtual HRESULT STDMETHODCALLTYPE FSetTrackingComponent(
		DWORD_PTR dwComponentID, BOOL fTrackBegin) = 0;
	virtual void STDMETHODCALLTYPE OnComponentEnterState(
		DWORD_PTR dwComponentID, Msocstate uStateID, MSOCCONTEXT uContext, ULONG cpicmExclude, 
		IMsoComponentManager **rgpicmExclude, DWORD dwReserved) = 0;
	virtual BOOL STDMETHODCALLTYPE FOnComponentExitState(
		DWORD_PTR dwComponentID, Msocstate uStateID, MSOCCONTEXT uContext, 
		ULONG cpicmExclude, IMsoComponentManager **rgpicmExclude ) = 0;
	virtual BOOL STDMETHODCALLTYPE FInState(Msocstate uStateID, void *pvoid/*reserved*/) = 0;
	virtual BOOL STDMETHODCALLTYPE FContinueIdle(void)= 0;
	virtual BOOL STDMETHODCALLTYPE FPushMessageLoop(
		DWORD_PTR dwComponentID, Msoloop uReason, void *pvLoopData)= 0;
	// always returns FALSE
	virtual BOOL STDMETHODCALLTYPE FCreateSubComponentManager(IUnknown *_opt_piunkOuter, 
		IUnknown *_opt_piunkServProv, REFIID riid, void **ppvObj) = 0;
	// always returns FALSE
	virtual BOOL STDMETHODCALLTYPE FGetParentComponentManager(IMsoComponentManager **ppicm) = 0;
	virtual BOOL STDMETHODCALLTYPE FGetActiveComponent(MSOGAC dwgac, 
		IMsoComponent **ppic, MSOCRINFO *pcrinfo, DWORD dwReserved)= 0;
};

#endif

