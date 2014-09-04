// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0501		// Allow use of features specific to Windows XP or later.
#endif						

#ifndef _WIN32_IE
#define _WIN32_IE 0x0800		// Allow use of features specific to IE 8.0 or later.
#endif

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

// turns off ATL's hiding of some common and often safely ignored warning messages
#define _ATL_ALL_WARNINGS

#include "resource.h"

#include <atlbase.h>
#include <atlcom.h>

#include "Import/MSO.tlh"
#include "Import/VISLIB.tlh"
#include "Import/IDTExtensibility2.tlh"

using namespace ATL;
extern CComModule _Module;
