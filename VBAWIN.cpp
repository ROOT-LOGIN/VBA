// VBAWIN.cpp : WinMain 的实现

#include "stdafx.h"
#include "resource.h"
#include "VBAWIN_I.H"
#include "VBAWIN.H"

#include "../Include/language_token.h"

#pragma comment(lib, "comctl32.lib")

const IID IID_IMsoComponent = {
	0x000C0600, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 }
};

const IID IID_IMsoComponentManager = {
	0x000C060B, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 }
};

const IID SID_SMsoComponentManager = {
	0x000C0601, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 }
};

const IID IID_VBE = {
	0x0002E166L,0x0000,0x0000, { 0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46 }
};

VBAWINMAINWIN *pmainwin = NULL;

class CVBAWINModule : 
	public ATL::CAtlExeModuleT< CVBAWINModule >
{
public :
	DECLARE_LIBID(LIBID_VBAWINLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_VBAWIN, "{968D9076-9436-46D1-A59C-0B8FF695F45C}")

	HINSTANCE m_hInstance;
};

CVBAWINModule _AtlModule;

//OBJECT_ENTRY_NON_CREATEABLE_EX_AUTO(CLSID_VBAWINApp, VBAWINApp);


#ifdef PURE
#undef PURE
#endif

#define PURE { return S_OK; }

// {115dd106-2b24-11d2-aa6c-00c04f9901d2}
const IID IID_IVbaSite = {
	0x115dd106, 0x2b24, 0x11d2, { 0xaa, 0x6c, 0x0, 0xc0, 0x4f, 0x99, 0x01, 0xd2 }
};

class VBASITE :
	public CComObjectRoot,
	public IVbaSite
{
	BEGIN_COM_MAP(VBASITE)
		COM_INTERFACE_ENTRY_IID(IID_IVbaSite, IVbaSite)
	END_COM_MAP()

public:
	// *** IVbaSite methods ***
	STDMETHOD(GetOwnerWindow)(THIS_ HWND *phwnd) {
		*phwnd = pmainwin->m_hWnd;
		return S_OK;
	}
	STDMETHOD(Notify)(THIS_ VBASITENOTIFY vbahnotify) {
		return S_OK;
	}

	STDMETHOD(FindFile)(THIS_ LPCOLESTR lpstrFileName, BSTR *pbstrFullPath) {
		return S_OK;
	}
	STDMETHOD(GetMiniBitmap)(THIS_ REFGUID rgguid, HBITMAP *phbmp, COLORREF *pcrMask) {
		if(rgguid == GUID_NULL)
		{
			*phbmp = LoadBitmap(_AtlModule.m_hInstance, MAKEINTRESOURCE(IDB_VABWINMAIN));
			*pcrMask = RGB(255,0,255);
		}
		return S_OK;
	}

	STDMETHOD(CanEnterBreakMode)(THIS_ VBA_BRK_INFO *pbrkinfo) {
		return S_OK;
	}
	STDMETHOD(SetStepMode)(THIS_ VBASTEPMODE vbastepmode) {
		return S_OK;
	}
	STDMETHOD(ShowHide)(THIS_ BOOL fVisible) {
		return S_OK;
	}

	STDMETHOD(InstanceCreated)(THIS_ IUnknown *punkInstance) {
		return S_OK;
	}

	STDMETHOD(HostCheckReference)(THIS_ BOOL fSave, REFGUID rgguid,
		UINT *puVerMaj, UINT *puVerMin) {
			return S_OK;
	}

	STDMETHOD(GetIEVersion)(THIS_  LONG *plMajVer, LONG *plMinVer, BOOL * pfCanInstall) {
		return E_NOTIMPL;
	}
	STDMETHOD(UseIEFeature)(THIS_  LONG  lMajVer,  LONG  lMinVer,  HWND hwndOwner) {
		return S_OK;
	}

	STDMETHOD(ShowHelp)(THIS_ HWND hwnd, LPCOLESTR szHelp, UINT wCmd, DWORD dwHelpContext, BOOL fWinHelp) {
		return S_OK;
	}

	STDMETHOD(OpenProject)(THIS_ HWND hwndOwner, IVbaProject *pVbaRefProj,
		LPCOLESTR lpstrFileName, IVbaProject **ppVbaProj) {
			return S_OK;
	}

};

// {B3205671-FE45-11cf-8D08-00A0C90F2732}
const IID IID_IVbaCompManagerSite = {
	0xb3205671, 0xfe45, 0x11cf, { 0x8d, 0x8, 0x0, 0xa0, 0xc9, 0xf, 0x27, 0x32 }
};

// {1B30B6B8-E44B-11d1-9915-00A0C9702442}
const IID IID_IVbaProjectSite = {
	0x1b30b6b8, 0xe44b, 0x11d1, { 0x99, 0x15, 0x0, 0xa0, 0xc9, 0x70, 0x24, 0x42 }
};

class VBAPROJECTSITE :
	public CComObjectRoot,
	public IVbaProjectSite
{
	BEGIN_COM_MAP(VBAPROJECTSITE)
		COM_INTERFACE_ENTRY_IID(IID_IVbaProjectSite, IVbaProjectSite)
	END_COM_MAP()

public:
	STDMETHOD(GetOwnerWindow)(THIS_ HWND *phwnd)
	{
		*phwnd = pmainwin->m_hWnd;
		return S_OK;
	}
	STDMETHOD(Notify)(THIS_ VBAPROJSITENOTIFY vbapsn) 
	{
		if(vbapsn == VBAPSN_Activate)
			pmainwin->BringWindowToTop();
		return S_OK;
	}
	STDMETHOD(GetObjectOfName)(THIS_ LPCOLESTR lpstrName, IDocumentSite **ppDocSite)
	{
		return S_OK;
	}
	STDMETHOD(LockProjectOwner)(THIS_ BOOL fLock, BOOL fLastUnlockCloses) 
	{
		return S_OK;
	}
	STDMETHOD(RequestSave)(THIS_ HWND hwndOwner){
		return S_OK;
	}

	STDMETHOD(ReleaseModule)(THIS_ IVbaProjItem *pVbaProjItem) {
		return S_OK;
	}
	STDMETHOD(ModuleChanged)(THIS_ IVbaProjItem *pVbaProjItem) {
		return S_OK;
	}

	STDMETHOD(DeletingProjItem)(THIS_ IVbaProjItem *pVbaProjItem) {
		return S_OK;
	}

	STDMETHOD(ModeChange)(THIS_ VBA_PROJECT_MODE pvbaprojmode) {
		return S_OK;
	}
	STDMETHOD(ProcChanged)(THIS_ VBA_PROC_CHANGE_INFO *pvbapcinfo) {
		return S_OK;
	}

	STDMETHOD(ReleaseInstances)(THIS_ IVbaProjItem *pVbaProjItem) {
		return S_OK;
	}

	STDMETHOD(ProjItemCreated)(THIS_ IVbaProjItem *pVbaProjItem) {
		WCHAR path[MAX_PATH];
		GetModuleFileNameW(NULL, path, MAX_PATH);
		LPWSTR exe = PathFindFileName(path);
		wmemset(exe, 0, wcslen(exe));
		PathCombineW(path, path, L"Resources\\main.vb");
		pVbaProjItem->InsertTextIntoModule(path);
		return S_OK;
	}

	STDMETHOD(GetMiniBitmapGuid)(THIS_ IVbaProjItem *pVbaProjItem, GUID *pguidMiniBitmap) {
		return S_OK;
	}

	STDMETHOD(GetSignature)(THIS_  void **ppvSignature, DWORD * pcbSign) {
		return S_OK;
	}
	STDMETHOD(PutSignature)(THIS_  void  *pvSignature,  DWORD   cbSign)  {
		return S_OK;
	}

	STDMETHOD(NameChanged)(THIS_ IVbaProjItem *pVbaProjItem, LPCOLESTR lpstrOldName) {
		return S_OK;
	}
	STDMETHOD(ModuleDirtyChanged)(THIS_ IVbaProjItem *pVbaProjItem, BOOL fDirty) {
		return S_OK;
	}

	STDMETHOD(CreateDocClassInstance)(THIS_ IVbaProjItem *pVbaProjItem, CLSID clsidDocClass,
		IUnknown *punkOuter, BOOL fIsPredeclared,
		IUnknown **ppunkInstance) {
			return S_OK;
	}
};

// { 94A0F6F1-10BC-11d0-8D09-00A0C90F2732 }
const IID IID_IDocumentSite = {
	0x94a0f6f1, 0x10bc, 0x11d0, { 0x8d, 0x09, 0x00, 0xa0, 0xc9, 0x0f, 0x27, 0x32 }
};

// { 6d5140c5-7436-11ce-8034-00aa006009fa }
const IID IID_ITrackSelection = {
	0x6d5140c5, 0x7436, 0x11ce, { 0x80, 0x34, 0x00, 0xaa, 0x00, 0x60, 0x09, 0xfa }
};

const IID IID_ISelectionContainer = {
	0x6d5140c6, 0x7436, 0x11ce, { 0x80, 0x34, 0x00, 0xaa, 0x00, 0x60, 0x09, 0xfa }
};

const IID IID_ICategorizeProperties = {
	0x4d07fc10, 0xf931, 0x11ce, { 0xb0, 0x1, 0x0, 0xaa, 0x0, 0x68, 0x84, 0xe5 }
};

const IID IID_IControlContainer = {
	0xbc863490, 0xe675, 0x11ce, { 0x82, 0x29, 0x00, 0xaa, 0x00, 0x44, 0x40, 0xd0 }
};


class VBAWINProj : 
	public CComObjectRoot,
	public IDispatchImpl<IVBAWINProj, &IID_IVBAWINProj, &LIBID_VBAWINLib>,
	public ICategorizeProperties,
	public IControlContainer
{
	DECLARE_NO_REGISTRY()

	BEGIN_COM_MAP(VBAWINProj)
		COM_INTERFACE_ENTRY_IID(IID_IVBAWINProj, IVBAWINProj)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
		COM_INTERFACE_ENTRY_IID(IID_IControlContainer, IControlContainer)
	END_COM_MAP()

public:
	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Application( 
		/* [retval][out] */ IVBAWINApp **ppApp)
	{
		return QueryInterface(IID_IVBAWINApp, (void**)ppApp);
	}

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Parent( 		
		/* [retval][out] */ IVBAWINApp **ppApp)
	{
		return QueryInterface(IID_IVBAWINApp, (void**)ppApp);
	}

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Version( 
		/* [retval][out] */ LPBSTR name)
	{
		*name = SysAllocString(L"2013.04.24.5842");
		return S_OK;
	}

	virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE Notify( 
		/* [in] */ long dwCode)
	{
		return S_OK;
	}

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Location( 
		/* [retval][out] */ BSTR *lpbstrPath)
	{
		path_buf[279] = L'\0';
		*lpbstrPath = SysAllocString(path_buf);
		return S_OK;
	}

	virtual /* [helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_Location( 
		/* [in] */ BSTR bstrPath)
	{
		wcsncpy_s(path_buf, 279, bstrPath, 279);
		path_buf[279] = L'\0';
		return S_OK;
	}

	virtual /* [helpstring][propget] */ HRESULT STDMETHODCALLTYPE get_Token( 
		/* [retval][out] */ BSTR *lpbstrToken)
	{
		token_buf[279] = L'\0';
		*lpbstrToken = SysAllocString(token_buf);
		return S_OK;
	}

	virtual /* [helpstring][propput] */ HRESULT STDMETHODCALLTYPE put_Token( 
		/* [in] */ BSTR bstrToken)
	{
		wcsncpy_s(token_buf, 279, bstrToken, 279);
		token_buf[279] = L'\0';
		return S_OK;
	}

	// *** ICategorizeProperties ***
	STDMETHOD(MapPropertyToCategory)(THIS_ 
		/* [in]  */ DISPID dispid,
		/* [out] */ PROPCAT* ppropcat) 
	{
		DISPID idTkn, idLoc;
		WCHAR *nmTkn = L"Token", *nmLoc = L"Location";
		GetIDsOfNames(GUID_NULL, &nmTkn, 1, 2052, &idTkn);
		GetIDsOfNames(GUID_NULL, &nmLoc, 1, 2052, &idLoc);
		__ifequ(dispid, idTkn)
			*ppropcat = PROPCAT_Behavior;
		__elifequ(dispid, idLoc)
			*ppropcat = PROPCAT_Data;
		else
			*ppropcat = 0;
		return S_OK;
	}

	STDMETHOD(GetCategoryName)(THIS_
		/* [in]  */ PROPCAT propcat, 
		/* [in]  */ LCID lcid,
		/* [out] */ BSTR* pbstrName)
	{
		switch(propcat)
		{
		case PROPCAT_Nil: *pbstrName = SysAllocString(L"常规"); break;
		case PROPCAT_Misc: *pbstrName = SysAllocString(L"杂项"); break;
		case PROPCAT_Font: *pbstrName = SysAllocString(L"字体"); break;
		case PROPCAT_Position: *pbstrName = SysAllocString(L"位置"); break;
		case PROPCAT_Appearance: *pbstrName = SysAllocString(L"外观"); break;
		case PROPCAT_Behavior: *pbstrName = SysAllocString(L"行为"); break;
		case PROPCAT_Data: *pbstrName = SysAllocString(L"数据"); break;
		case PROPCAT_List: *pbstrName = SysAllocString(L"列表"); break;
		case PROPCAT_Text: *pbstrName = SysAllocString(L"文本"); break;
		case PROPCAT_Scale: *pbstrName = SysAllocString(L"缩放"); break;
		case PROPCAT_DDE: *pbstrName = SysAllocString(L"DDE"); break;
		default: *pbstrName = SysAllocString(L"未定义"); break;
		}
		return S_OK;
	}

	// *** IControlContainer methods ***
	STDMETHOD(Init)(THIS_ IControlRT *pctrlRT) {
		if(pctrlRT) pctrlRT->AddRef();
		mp_CtrlRT = pctrlRT;
		return S_OK;
	}

	STDMETHOD(GetControlParts)(THIS_ 
		DWORD dwflags,                
		DWORD dwCtlCookie,            
		IUnknown *punkOuter,          
		DWORD    *pdwRTCookie,
		IUnknown **ppunkExtender,     
		IUnknown **ppunkprivControl,  
		IOleControlSite **ppocs){
			return E_NOTIMPL;
	}

	STDMETHOD(CopyControlParts)(THIS_ 
		DWORD dwflags,               
		DWORD dwCtlCookieTemplate,   
		IUnknown *punkOuter,         
		DWORD *pdwCltCookie,         
		DWORD *pdwRTCookie,         
		IUnknown **ppunkExtender,    
		IUnknown **ppunkprivControl, 
		IOleControlSite **ppocs){
			return E_NOTIMPL;
	}

	STDMETHOD(SetExtenderEventSink)(THIS_ 
		DWORD dwCtlCookie,      
		IUnknown *punkSink,
		IPropertyNotifySink *ppropns) {
			return S_OK;
	}

	IControlRT *mp_CtrlRT;

private:
	OLECHAR path_buf[280];
	OLECHAR token_buf[280];
};

const IID IID_IDocumentSite2 = {
	0x61d4a8a1, 0x2c90, 0x11d2, { 0xad, 0xe4, 0x0, 0xc0, 0x4f, 0x98, 0xf4, 0x17 }
};

const IID IID_IExtendedTypeLib = {
	0x6d5140d6, 0x7436, 0x11ce, { 0x80, 0x34, 0x00, 0xaa, 0x00, 0x60, 0x09, 0xfa }
};

class VBAWINDOCSITE :
	public CComObjectRoot,
	public ISelectionContainer,
	public IDocumentSite2,
	public IControlContainer,	
	public IProvideClassInfo2Impl<&CLSID_VBAWINProj, &IID_IVBAWINProj, &LIBID_VBAWINLib>
{
	//DECLARE_CLASSFACTORY()
	//DECLARE_REGISTRY_RESOURCEID(IDR_VBAWIN)
	DECLARE_NO_REGISTRY()

	BEGIN_COM_MAP(VBAWINDOCSITE)
		COM_INTERFACE_ENTRY(IProvideClassInfo2)
		COM_INTERFACE_ENTRY(IProvideClassInfo)
		COM_INTERFACE_ENTRY_IID(IID_ISelectionContainer, ISelectionContainer)
		COM_INTERFACE_ENTRY_IID(IID_IControlContainer, IControlContainer)
		COM_INTERFACE_ENTRY_IID(IID_IDocumentSite2, IDocumentSite2)
		COM_INTERFACE_ENTRY_IID(IID_IDocumentSite, IDocumentSite)		
	END_COM_MAP()

public:
	VBAWINDOCSITE(void) {
		mp_ServiceProvider = NULL;
		mp_TrackSelection = NULL;
	}
	~VBAWINDOCSITE(void) {
		//if(mp_ServiceProvider) mp_ServiceProvider->Release();
		//if(mp_TrackSelection) mp_TrackSelection->Release();
	}

	// *** IDocumentSite methods ***
	STDMETHOD(SetSite)(THIS_/* [in]  */ IServiceProvider * pSP) {
		mp_ServiceProvider = pSP;
		pSP->AddRef();
		mp_ServiceProvider->QueryService(SID_STrackSelection, IID_ITrackSelection, (void**)&mp_TrackSelection);
		m_Proj.AddRef();
		mp_TrackSelection->OnSelectChange(this);
		CComPtr<IExtendedTypeLib> extplib;
		if(SUCCEEDED(mp_ServiceProvider->QueryService(SID_SExtendedTypeLib, IID_IExtendedTypeLib, (LPVOID*)&extplib))){
			CComQIPtr<IProvideClassInfo> pci = this;	
			if(pci){
				CComPtr<ITypeInfo> extinfo;
				pci->GetClassInfoW(&extinfo);
				if(extinfo)
					extplib->SetExtenderInfo(NULL, extinfo, NULL);
			}
		}
		
		return S_OK;
	}
	STDMETHOD(GetSite)(THIS_/* [out] */ IServiceProvider** ppSP) {
		*ppSP = mp_ServiceProvider;
		mp_ServiceProvider->AddRef();
		return S_OK;
	}
	STDMETHOD(GetCompiler)(THIS_/* [in]  */ REFIID iid,/* [out] */ void **ppvObj) {
		return QueryInterface(iid, ppvObj);
	}
	STDMETHOD(ActivateObject)(THIS_ DWORD dwFlags) {
		return S_OK;
	}
	STDMETHOD(IsObjectShowable)(THIS) {
		return S_OK;
	}

	// *** ISelectionContainer methods ***
	STDMETHOD(CountObjects)(THIS_/* [in]  */ DWORD dwFlags, /* [out] */ ULONG * pc) {
		*pc = 1;
		return S_OK;
	}
	STDMETHOD(GetObjects)(THIS_/* [in]  */ DWORD dwFlags,
		/* [in]  */ ULONG cObjects, /* [out] */ IUnknown **apUnkObjects) {
			cObjects = 1;
			m_Proj.QueryInterface(apUnkObjects);			
			return S_OK;
	}
	STDMETHOD(SelectObjects)(THIS_/* [in] */ ULONG cSelect,
		/* [in] */ IUnknown **apUnkSelect, /* [in] */ DWORD dwFlags) {
			m_Proj.QueryInterface(apUnkSelect);
			return S_OK;
	}

	// *** IDocumentSite2 methods ***
	STDMETHOD(GetObject)(THIS_
		/* [out] */ IDispatch** ppdispObj)
	{
		return m_Proj.QueryInterface(IID_IDispatch, (void**)ppdispObj);
	}

	
	// *** IControlContainer methods ***
	STDMETHOD(Init)(THIS_ IControlRT *pctrlRT) {
		if(pctrlRT) pctrlRT->AddRef();
		mp_CtrlRT = pctrlRT;
		return S_OK;
	}

	STDMETHOD(GetControlParts)(THIS_ 
		DWORD dwflags,                
		DWORD dwCtlCookie,            
		IUnknown *punkOuter,          
		DWORD    *pdwRTCookie,
		IUnknown **ppunkExtender,     
		IUnknown **ppunkprivControl,  
		IOleControlSite **ppocs){
			return E_NOTIMPL;
	}

	STDMETHOD(CopyControlParts)(THIS_ 
		DWORD dwflags,               
		DWORD dwCtlCookieTemplate,   
		IUnknown *punkOuter,         
		DWORD *pdwCltCookie,         
		DWORD *pdwRTCookie,         
		IUnknown **ppunkExtender,    
		IUnknown **ppunkprivControl, 
		IOleControlSite **ppocs){
			return E_NOTIMPL;
	}

	STDMETHOD(SetExtenderEventSink)(THIS_ 
		DWORD dwCtlCookie,      
		IUnknown *punkSink,
		IPropertyNotifySink *ppropns) {
			return S_OK;
	}

	CComObject<VBAWINProj> m_Proj;
private:
	IServiceProvider *mp_ServiceProvider;
	ITrackSelection *mp_TrackSelection;	
	IControlRT *mp_CtrlRT;
};

// { 759d0500-d979-11ce-84ec-00aa00614f3e }
const IID IID_IUIElement = {
	0x759d0500, 0xd979, 0x11ce, { 0x84, 0xec, 0x00, 0xaa, 0x00, 0x61, 0x4f, 0x3e }
};

MODULEHANDLE* gp_hMso = NULL;
const IID* _IID_REFS[30];
int sjajs()
{
#define COMBINE(l, r) l ## r
#define NAMEIS(name, i) COMBINE(_Interface_,i); _IID_REFS[i] = &__uuidof(COMBINE(_Interface_,i))
struct __declspec(uuid("0002e158-0000-0000-c000-000000000046")) NAMEIS(Application, __COUNTER__);
struct __declspec(uuid("0002e166-0000-0000-c000-000000000046")) NAMEIS(VBE, __COUNTER__);
struct __declspec(uuid("0002e16b-0000-0000-c000-000000000046")) NAMEIS(Window, __COUNTER__);
struct __declspec(uuid("0002e16a-0000-0000-c000-000000000046")) NAMEIS(_Windows_old, __COUNTER__);
struct __declspec(uuid("f57b7ed0-d8ab-11d1-85df-00c04f98f42c")) NAMEIS(_Windows, __COUNTER__);
struct __declspec(uuid("0002e16c-0000-0000-c000-000000000046")) NAMEIS(_LinkedWindows, __COUNTER__);
struct __declspec(uuid("0002e159-0000-0000-c000-000000000046")) NAMEIS(_ProjectTemplate, __COUNTER__);
struct __declspec(uuid("0002e160-0000-0000-c000-000000000046")) NAMEIS(_VBProject_Old, __COUNTER__);
struct __declspec(uuid("eee00915-e393-11d1-bb03-00c04fb6c4a6")) NAMEIS(_VBProject, __COUNTER__);
struct __declspec(uuid("0002e165-0000-0000-c000-000000000046")) NAMEIS(_VBProjects_Old, __COUNTER__);
struct __declspec(uuid("eee00919-e393-11d1-bb03-00c04fb6c4a6")) NAMEIS(_VBProjects, __COUNTER__);
struct __declspec(uuid("be39f3d4-1b13-11d0-887f-00a0c90f2744")) NAMEIS(SelectedComponents, __COUNTER__);
struct __declspec(uuid("0002e161-0000-0000-c000-000000000046")) NAMEIS(_Components, __COUNTER__);
struct __declspec(uuid("0002e162-0000-0000-c000-000000000046")) NAMEIS(_VBComponents_Old, __COUNTER__);
struct __declspec(uuid("eee0091c-e393-11d1-bb03-00c04fb6c4a6")) NAMEIS(_VBComponents, __COUNTER__);
struct __declspec(uuid("0002e163-0000-0000-c000-000000000046")) NAMEIS(_Component, __COUNTER__);
struct __declspec(uuid("0002e164-0000-0000-c000-000000000046")) NAMEIS(_VBComponent_Old, __COUNTER__);
struct __declspec(uuid("eee00921-e393-11d1-bb03-00c04fb6c4a6")) NAMEIS(_VBComponent, __COUNTER__);
struct __declspec(uuid("0002e18c-0000-0000-c000-000000000046")) NAMEIS(Property, __COUNTER__);
struct __declspec(uuid("0002e188-0000-0000-c000-000000000046")) NAMEIS(_Properties, __COUNTER__);
struct __declspec(uuid("da936b62-ac8b-11d1-b6e5-00a0c90f2744")) NAMEIS(_AddIns, __COUNTER__);
struct __declspec(uuid("da936b64-ac8b-11d1-b6e5-00a0c90f2744")) NAMEIS(AddIn, __COUNTER__);
struct __declspec(uuid("0002e16e-0000-0000-c000-000000000046")) NAMEIS(_CodeModule, __COUNTER__);
struct __declspec(uuid("0002e172-0000-0000-c000-000000000046")) NAMEIS(_CodePanes, __COUNTER__);
struct __declspec(uuid("0002e176-0000-0000-c000-000000000046")) NAMEIS(_CodePane, __COUNTER__);
struct __declspec(uuid("0002e17a-0000-0000-c000-000000000046")) NAMEIS(_References, __COUNTER__);
struct __declspec(uuid("0002e17e-0000-0000-c000-000000000046")) NAMEIS(Reference, __COUNTER__);
return __COUNTER__;
}

extern "C" const IID SID_SVbaMainUI;/* = {
	0x759d0501, 0xd979, 0x11ce, { 0x84,0xec,0x00,0xaa,0x00,0x61,0x4f,0x3e }
};*/

const IID IID_IControlDesign = {
	0x5bd18670, 0xe2fa, 0x11ce, { 0x82, 0x29, 0x00, 0xaa, 0x00, 0x44, 0x40, 0xd0 }
};

const IID CLSID_VBAWINActiveDesigner = {
	0x2A04E498, 0xD5F4, 0x4792, { 0xA1, 0x24, 0x1F, 0x9E, 0x25, 0xD0, 0xAF, 0x8E }
};

extern "C" int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, 
								LPTSTR lpCmdLine, int nShowCmd)
{	
#if (defined(DEBUG) || defined(_DEBUG)) && !defined(NO_DBG_PRINTF)
	AllocConsole();
	_wfreopen(L"CONOUT$", L"r+", stdout);
#endif
	InitCommonControls();

	int d = sjajs();

	HRESULT hr = S_OK;
	//if(!_AtlModule.ParseCommandLine(lpCmdLine, &hr))
		//return hr;
	_AtlModule.m_hInstance = hInstance;
	VBAWINMAINWIN mainwin;

	CComObject<VBASITE> vbasite;
	CComObject<VBACOMPMGRSITE> vbacompmgrsite;
	CComObject<VBAPROJECTSITE> vbaprojsite;
	CComObject<VBAWINDOCSITE> vbadocsite;
	vbaprojsite.AddRef();
	vbadocsite.AddRef();

	CRect rc(0,0,360,240);	
	mainwin.Create(NULL, _U_RECT(rc), L"VBAWIN MAIN WINDOW", 0, 0, _U_MENUorID(LoadMenu(hInstance, MAKEINTRESOURCE(IDR_MAINMENU))));
	pmainwin = &mainwin;

	VBASAPROJINFO proj;
	ZeroMemory(&proj, sizeof(proj));
	proj.lpstrHostApp = L"VBAWIN MAIN APPLICATION";
	proj.lpstrRegPath = L"Software\\VBAWIN";
	proj.vaodHost.uVerMaj = 1;
	proj.vaodHost.uVerMin = 0;
	proj.vaodHost.guidTypelibAppObj = LIBID_VBAWINLib;
	proj.vsaptType = VBASAPT_NONE;

	VBAINITINFO info;
	ZeroMemory(&info, sizeof(info));
	info.cbSize = sizeof(info);
	vbasite.QueryInterface(IID_IVbaSite, (void**)&info.pVbaSite);
	info.lpstrName = L"VBAWIN MAIN APPLICATION";
	info.lcidUserInterface = 2052;
	info.lpstrProjectFileFilter = L"All Reference Files (*.olb, *.tlb, *.dll, *.exe, *.ocx)\0*.olb;*.tlb;*.dll;*.exe;*.ocx\0";
	mainwin.m_AppObject.QueryInterface(IID_IUnknown, (LPVOID*)&info.punkAppObject);
	info.hwndHost = mainwin;
	info.hinstHost = hInstance;
	vbacompmgrsite.QueryInterface(IID_IVbaCompManagerSite, (void**)&info.pVbaCompManagerSite);
	info.dwFlags = VBAINITINFO_MessageLoopWide | VBAINITINFO_SupportVbaDesigners;
	info.cSAProjTypes = 0;
	info.rgvsapiProjTypes = NULL;//&proj;
	info.vaodHost.uVerMaj = 1;
	info.vaodHost.uVerMin = 0;
	info.vaodHost.guidTypelibAppObj = LIBID_VBAWINLib;
	info.vaodHost.guidTypeinfoAppObj = CLSID_VBAWINApp;
	
	VBAINITHOSTADDININFO addin;
	addin.dwCookie = 'SAYU';
	WCHAR pth[MAX_PATH];
	wcscpy(pth, L"Software\\VBAWIN\\AddIns");
	addin.lpstrRegPath = pth;
	addin.lcidHost = 2052;
	addin.dwFlags = 0;
	addin.fRemoveFromListOnFailure = FALSE;
	mainwin.m_AppObject.QueryInterface(IID_IDispatch, (void**)&addin.pdisp);

#if !defined(_HOST_X64) || !_HOST_X64
	MODULEHANDLE hMso(_T("MSO.DLL"));
	info.hinstReserved = hMso;
	gp_hMso = &hMso;
	MODULEHANDLE hClbcatq(_T("clbcatq.dll"));
	MODULEHANDLE hShl32(_T("shell32.dll"));
	/*MODULEHANDLE hNetapi(_T("netapi32.dll"));
	MODULEHANDLE hBetuti(_T("netutils.dll"));
	MODULEHANDLE hPropsys(_T("propsys.dll"));
	MODULEHANDLE hProfapi(_T("profapi.dll"));
	MODULEHANDLE hSrvcli(_T("srvcli.dll"));
	MODULEHANDLE hWkscli(_T("wkscli.dll"));
	MODULEHANDLE hRpcrem(_T("RpcRtRemote.dll"));*/
#endif

	if(1)//mainwin.m_AppObject.m_hVbe)
	{
#if !defined(_HOST_X64) || !_HOST_X64
		DoDetour(NULL);//mainwin.m_AppObject.m_hVbe);
#endif
		CComPtr<IVbaProject> proj;
		CComPtr<IStorage> stg;
		hr = DllVbeInit(6, 0, 0x653, 0, 0x0FF1CE, &info, &mainwin.mp_Vba);
		hr = DllVbaInitHostAddins(&addin, &mainwin.mp_HostAddins);

		mainwin.ShowWindow(SW_SHOW);
			
		hr = mainwin.mp_Vba->CreateProject(VBAPROJFLAG_None, &vbaprojsite, &proj);
		hr = StgCreateDocfile(NULL, STGM_TRANSACTED | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_CREATE | STGM_DELETEONRELEASE, NULL, &stg);
			
		if(proj && stg)
		{
			CComQIPtr<IPersistStorage> psstg = proj;
			psstg->InitNew(stg);
			((IUnknown*)proj)->AddRef();

			proj->SetName(L"VBAWINProj");
			proj->SetDisplayName(L"VBAWIN 项目");
			proj->SetMode(VBA_PROJ_MODE_Use);
			proj->AddTypeLibRef(LIBID_VBAWINLib, 1, 0);
				
			IUnknown* vbeEO;
			hr = mainwin.mp_Vba->CreateExtension(&mainwin.m_AppObject, (IVBAWINProj*)&vbadocsite.m_Proj, NULL, &vbeEO);

			CComPtr<IVbaProjItemCol> projitemcol;			
			hr = proj->GetProjItemCol(&projitemcol);
			if(projitemcol){
				//projitemcol->CreateModuleProjItem(L"ModuleProjItem", VBAMODTYPE_StdMod, NULL);
				CComPtr<IVbaProjItem> projitem;
				hr = projitemcol->CreateDocumentProjItem(L"DocumentProjItem", &vbadocsite, NULL, &projitem);
				if(projitem){
					//projitem->LoadDocClass();
					projitem->SetExtension(vbeEO);
					projitem->SetName(L"VBAWINDocumentProjItem");
					projitem->SetDisplayName(L"VBAWIN 文档项目");
					CComPtr<IControlDesign> ctrldes;
					CComQIPtr<IServiceProvider> projitmsrv = projitem;
					projitmsrv->QueryService(SID_SControlDesign, IID_IControlDesign, (LPVOID*)&ctrldes);
				}
#if 0
				projitem.Release();
				/* A document class is a class can be referenced directly in other project item. 
				 * that is it does't need do declare(Dim doc_cls As New DocClassProjectItem)*/
				hr = projitemcol->CreateDocClassProjItem(L"DocClassProjItem", &vbadocsite, &projitem);
				if(projitem){
					projitem->SetExtension(vbeEO);
					projitem->LoadDocClass();						
					projitem->SetName(L"VBAWINDocClassProjItem");
					projitem->SetDisplayName(L"VBAWIN 文档类项目");
				}
#endif
#if 1
				projitem.Release();
				hr = projitemcol->CreateDesignerProjItem(L"DesignerProjItem", CLSID_VBAWINActiveDesigner, &projitem);
				if(projitem){
					//projitem->SetExtension(vbeEO);
					//projitem->LoadDocClass();
					projitem->SetName(L"VBAWINDesignerProjItem");
					projitem->SetDisplayName(L"VBAWIN 设计器项目");
				}
#endif
				proj->LoadComplete();
			}
		}
		hr = mainwin.mp_HostAddins->LoadHostAddins();
		mainwin.mp_HostAddins->OnHostStartupComplete();
		/*CComQIPtr<IServiceProvider> srvprov = mainwin.mp_Vba;
		CComPtr<IUIElement> vbaui;			
		srvprov->QueryService(SID_SVbaMainUI, IID_IUIElement, (void**)&vbaui);
		vbaui->Show();
		vbaui->Hide();*/
		mainwin.RunMessageLoop();
		mainwin.mp_HostAddins->OnHostBeginShutdown();
		proj->Close();			
		DllVbeTerm(mainwin.mp_Vba);
	}
#if (defined(DEBUG) || defined(_DEBUG)) && !defined(NO_DBG_PRINTF)
	FreeConsole();
#endif
	return 0;
}
