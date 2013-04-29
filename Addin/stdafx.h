// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0600
#endif						

#ifndef _WIN32_IE
#define _WIN32_IE 0x0700
#endif

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit
#define ISOLATION_AWARE_ENABLED 1
#define _ATL_ALL_WARNINGS

#include "resource.h"

#include <afxwin.h>
#include <atlbase.h>
#include <atlcom.h>

#pragma warning( disable : 4278 )
#pragma warning( disable : 4146 )

#include "import\MSADDNDR.tlh"
#include "import\MSO.tlh"
#include "import\VISLIB.tlh"

#pragma warning( default : 4146 )
#pragma warning( default : 4278 )

using namespace ATL;
