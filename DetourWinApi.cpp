#include "stdafx.h"
#include "../Include/detours.h"
#if defined(UNICODE) || defined(_UNICODE)
#pragma comment(lib, "../Debug/detoursw.lib")
#else
#pragma comment(lib, "../Debug/detoursa.lib")
#endif

#ifdef ISOLATION_AWARE_ENABLED
#error DetourWinApp.cpp is incompatible with ISOLATION_AWARE_ENABLED
#endif

#define STREQU(l, r) (strcmp(l, r) == 0)
#define STRNEQU(l, r, n) (strncmp(l, r, n) == 0)
#define STRIEQU(l, r) (stricmp(l, r) == 0)
#define STRNIEQU(l, r, n) (strnicmp(l, r, n) == 0)
#define WCSEQU(l, r) (wcscmp(l, r) == 0)
#define WCSNEQU(l, r, n) (wcsncmp(l, r, n) == 0)
#define WCSIEQU(l, r) (_wcsicmp(l, r) == 0)
#define WCSNIEQU(l, r, n) (wcsnicmp(l, r, n) == 0)

#define FUNCNAME(ApiName) Dwa##ApiName
#define FUNCPTRTYPE(ApiName) Dwafnptr##ApiName
#define FUNCPTR(ApiName) ApiName##Ptr
#define SXSFUNCPTR(ApiName, sxs) ApiName_##sxs##_Ptr

#define DETOUR_FUNCPTR_DECL(ApiName) \
	FUNCPTRTYPE(ApiName) FUNCPTR(ApiName) = NULL;
#define DETOUR_FUNCPTR_DECL_SXS(ApiName, sxs) \
	FUNCPTRTYPE(ApiName) SXSFUNCPTR(ApiName, sxs) = NULL;

#define DEFINE_DWAAPI(RetType, ApiName, ...) \
	RetType WINAPI FUNCNAME(ApiName)(__VA_ARGS__)

#define DEFINE_DETOUR_OBJECT_NOPTR(RetType, ApiName, ...) \
	typedef RetType (WINAPI *FUNCPTRTYPE(ApiName))(__VA_ARGS__); \
	DEFINE_DWAAPI(RetType, ApiName, __VA_ARGS__);

#define DEFINE_DETOUR_OBJECT(RetType, ApiName, ...) \
	DEFINE_DETOUR_OBJECT_NOPTR(RetType, ApiName, __VA_ARGS__) \
	DETOUR_FUNCPTR_DECL(ApiName)

#define DETOUR_DO_ATTACH(ApiName) \
	FUNCPTR(ApiName) = ApiName; \
	DetourAttach((PVOID*)&FUNCPTR(ApiName), FUNCNAME(ApiName));

#define DETOUR_DO_ATTACH_EX(ApiName, InitVal) \
	FUNCPTR(ApiName) = (FUNCPTRTYPE(ApiName))InitVal; \
	DetourAttach((PVOID*)&FUNCPTR(ApiName), FUNCNAME(ApiName));

#define DETOUR_DO_ATTACH_BYNAME(ApiName, ModuleHandle, ProcName) \
	DETOUR_DO_ATTACH_EX(ApiName, GetProcAddress(ModuleHandle, #ProcName));
	
#define DETOUR_DO_ATTACH_BYNAMES(ApiName, ModuleHandle) \
	DETOUR_DO_ATTACH_EX(ApiName, GetProcAddress(ModuleHandle, #ApiName));

#define DETOUR_DO_DETACH(ApiName) \
	DetourDetach((PVOID*)&FUNCPTR(ApiName), FUNCNAME(ApiName));

// ==============================================
// DETOUR VBE API 
// ==============================================

DEFINE_DETOUR_OBJECT(void, rtcRandomize, VARIANT* _r8_or_empty)
DEFINE_DETOUR_OBJECT(double, rtcRandomNext, VARIANT* _r8_or_empty)

void DoDetourVbaRuntime(HMODULE hVbeDll)
{
	DETOUR_DO_ATTACH_BYNAMES(rtcRandomize, hVbeDll)
	DETOUR_DO_ATTACH_BYNAMES(rtcRandomNext, hVbeDll)
}

void WINAPI DwartcRandomize(VARIANT* _r8_or_empty)
{
	return;
}

double WINAPI DwartcRandomNext(VARIANT* _r8_or_empty)
{
	return 0.1;
}

// ==============================================
// DETOUR WIN API
// ==============================================

DEFINE_DETOUR_OBJECT(FARPROC, GetProcAddress, HMODULE hModule, LPCSTR procedure)
DEFINE_DETOUR_OBJECT(HMODULE, LoadLibraryA, LPCSTR library)
DEFINE_DETOUR_OBJECT(HMODULE, LoadLibraryW, LPCWSTR library)
DEFINE_DETOUR_OBJECT(HMODULE, LoadLibraryExA, LPCSTR lpFileName, HANDLE, DWORD dwFlags)
DEFINE_DETOUR_OBJECT(HMODULE, LoadLibraryExW, LPCWSTR lpFileName, HANDLE, DWORD dwFlags)
DEFINE_DETOUR_OBJECT(DWORD, MsiProvideQualifiedComponentA, \
	LPCSTR compcode, LPCSTR qualifier, DWORD installmode, LPSTR path, int *chpath)

DEFINE_DETOUR_OBJECT(HWND, CreateDialogParamA,  HINSTANCE hInstance, \
	LPCSTR lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
DEFINE_DETOUR_OBJECT(HWND, CreateDialogParamW,  HINSTANCE hInstance, \
	LPCWSTR lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
DEFINE_DETOUR_OBJECT(HWND, CreateDialogIndirectParamA, HINSTANCE hInstance, \
	LPCDLGTEMPLATEA lpTemplate, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM lParamInit)
DEFINE_DETOUR_OBJECT(HWND, CreateDialogIndirectParamW, HINSTANCE hInstance, \
	LPCDLGTEMPLATEW lpTemplate, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM lParamInit)

void DoDetour(HMODULE hVbeDll)
{
	GetProcAddressPtr = GetProcAddress;
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	//DETOUR_DO_ATTACH(LoadLibraryW);
	//DETOUR_DO_ATTACH(LoadLibraryA);
	//DETOUR_DO_ATTACH(LoadLibraryExW);
	//DETOUR_DO_ATTACH(LoadLibraryExA);
	//DETOUR_DO_ATTACH(CreateThread)
	//DETOUR_DO_ATTACH(CreateDialogParamA)
	//DETOUR_DO_ATTACH(CreateDialogParamW)
	//DETOUR_DO_ATTACH(CreateDialogIndirectParamA)
	//DETOUR_DO_ATTACH(CreateDialogIndirectParamW)
	//if(hVbeDll) \
		DoDetourVbaRuntime(hVbeDll);
	DETOUR_DO_ATTACH(GetProcAddress);
	DetourTransactionCommit();
}

void UndoDetour(void)
{
}

// =================================================
int WINAPI DwaMsiReinstallFeatureA(LPCSTR code, LPCSTR feature, DWORD mode);

bool strisinteger(LPSTR str)
{
	if(*str == '0')
	{
		str++;
		if(*str == 'x' || *str == 'X') // hex
		{
			str++;
			while(*str)
			{
				if(*str > '9' || *str < '0')
				{
					if(*str > 'f' || *str < 'a')
					{
						if(*str > 'F' || *str < 'A')
							return false;
					}
				}
				str++;
			}
		}
		else // oct
		{
			while(*str)
			{
				if(*str > '7' || *str < '0')
					return false;
				str++;
			}
			return true;
		}
	}
	else // dec
	{
		while(*str)
		{
			if(*str > '9' || *str < '0')
				return false;
			str++;
		}
		return true;

	}
	return false;
}

FARPROC WINAPI DwaGetProcAddress(HMODULE hModule, LPCSTR procedure)
{
	if((((DWORD)procedure) & 0xFFFF0000) == 0)
		return GetProcAddressPtr(hModule, procedure);

	if STREQU(procedure, "MsiReinstallFeatureA")
	{
		return (FARPROC)DwaMsiReinstallFeatureA; 
	}
	else if STREQU(procedure, "MsiProvideQualifiedComponentA")
	{
		MsiProvideQualifiedComponentAPtr = (DwafnptrMsiProvideQualifiedComponentA)GetProcAddressPtr(hModule, procedure);
		return (FARPROC)DwaMsiProvideQualifiedComponentA;
	}
	else
		return GetProcAddressPtr(hModule, procedure);
}

DWORD WINAPI DwaMsiProvideQualifiedComponentA(
	LPCSTR compcode, LPCSTR qualifier, DWORD installmode, LPSTR path, int *chpath)
{	
	DWORD _ret = MsiProvideQualifiedComponentAPtr(compcode,qualifier,installmode,path,chpath);
	if(_ret != NOERROR)
	{
		strcpy(path, "MSO.DLL");
		*chpath = sizeof("MSO.DLL")-1;
		return NOERROR;
	}
	return _ret;
}

int WINAPI DwaMsiReinstallFeatureA(LPCSTR code, LPCSTR feature, DWORD mode)
{
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());
	DETOUR_DO_DETACH(GetProcAddress);
	DetourTransactionCommit();
	return S_OK;
}

HMODULE WINAPI DwaLoadLibraryA(LPCSTR library)
{
	dbgprintf("LoadLibraryA #%d %s\n", GetCurrentThreadId(), library);
	return LoadLibraryAPtr(library);
}

HMODULE WINAPI DwaLoadLibraryW(LPCWSTR library)
{
	wdbgprintf(L"LoadLibraryW #%d %s\n", GetCurrentThreadId(), library);
	return LoadLibraryWPtr(library);
}

HMODULE WINAPI DwaLoadLibraryExA(LPCSTR lpFileName, HANDLE hFile, DWORD dwFlags)
{
	dbgprintf("LoadLibraryExA #%d %s\n", GetCurrentThreadId(), lpFileName);
	return LoadLibraryExAPtr(lpFileName, hFile, dwFlags);
}

extern MODULEHANDLE* gp_hMso;

HMODULE WINAPI DwaLoadLibraryExW(LPCWSTR lpFileName, HANDLE hFile, DWORD dwFlags)
{
	wdbgprintf(L"LoadLibraryExW #%d %s\n", GetCurrentThreadId(), lpFileName);
	LPWSTR library = PathFindFileNameW(lpFileName);	
	if WCSIEQU(library, L"MSORES.DLL")
		lpFileName = library;
	else if WCSIEQU(library, L"OGL.DLL")
		lpFileName = library;
	else if WCSIEQU(library, L"MSXML5.DLL")
		lpFileName = library;
	else if WCSIEQU(library, L"MSO.DLL")
		return (*gp_hMso);
	/*else if WCSIEQU(library, L"OSETUPPS.DLL")
		return NULL;
	else if WCSIEQU(library, L"sfc_os.dll")
		return NULL;
	else if WCSIEQU(library, L"sfc.dll")
		return NULL;
	else if WCSIEQU(library, L"setupapi.dll")
		return NULL;*/
	return LoadLibraryExWPtr(lpFileName, hFile, dwFlags);
}

DEFINE_DWAAPI(HWND, CreateDialogParamA,  HINSTANCE hInstance, \
	LPCSTR lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
	return FUNCPTR(CreateDialogParamA)(hInstance, lpTemplateName, hWndParent, lpDialogFunc, dwInitParam);
}

DEFINE_DWAAPI(HWND, CreateDialogParamW,  HINSTANCE hInstance, \
	LPCWSTR lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
	return FUNCPTR(CreateDialogParamW)(hInstance, lpTemplateName, hWndParent, lpDialogFunc, dwInitParam);
}

DEFINE_DWAAPI(HWND, CreateDialogIndirectParamA, HINSTANCE hInstance, \
	LPCDLGTEMPLATEA lpTemplate, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM lParamInit)
{
	return FUNCPTR(CreateDialogIndirectParamA)(hInstance, lpTemplate, hWndParent, lpDialogFunc, lParamInit);
}

DEFINE_DWAAPI(HWND, CreateDialogIndirectParamW, HINSTANCE hInstance, \
	LPCDLGTEMPLATEW lpTemplate, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM lParamInit)
{
	return FUNCPTR(CreateDialogIndirectParamW)(hInstance, lpTemplate, hWndParent, lpDialogFunc, lParamInit);
}
