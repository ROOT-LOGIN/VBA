// stdafx.h : 标准系统包含文件的包含文件，
// 或是经常使用但不常更改的
// 特定于项目的包含文件

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

//#define ISOLATION_AWARE_ENABLED 1

#define _CRT_SECURE_NO_WARNINGS

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// 某些 CString 构造函数将是显式的


#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>

/*#ifdef RGB
#pragma push_macro("RGB")
#undef RGB
#define __VBA_RGB_PUSHED
#endif
#ifdef EOF
#pragma push_macro("EOF")
#undef EOF
#define __VBA_EOF_PUSHED
#endif
#ifdef GetObject
#pragma push_macro("GetObject")
#undef GetObject
#define __VBA_GETOBJECT_PUSHED
#endif*/

//#import "C:/Program Files (x86)/Common Files/Microsoft Shared/OFFICE14/MSO.DLL"
//#import "C:/Program Files (x86)/Common Files/microsoft shared/VBA/VBA6/VBE6EXT.OLB" raw_interfaces_only
//#import "../debug/vbe7.dll" raw_interfaces_only

/*#ifdef __VBA_GETOBJECT_PUSHED
#pragma pop_macro("GetObject")
#endif
#ifdef __VBA_EOF_PUSHED
#pragma pop_macro("EOF")
#endif
#ifdef __VBA_RGB_PUSHED
#pragma pop_macro("RGB")
#endif*/

#include <atlctl.h>
#include <atltypes.h>

using namespace ATL;

#include "../VBA/VBA.H"
