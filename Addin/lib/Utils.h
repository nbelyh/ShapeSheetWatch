
#pragma once

/**----------------------------------------------------------------------
	Macro to be used on each COM method entry - try + AFX_MANAGE_STATE.
------------------------------------------------------------------------*/

#define ENTER_METHOD()								\
	try												\
	{												\
		AFX_MANAGE_STATE(AfxGetAppModuleState());	\

/**----------------------------------------------------------------------
	Macro to be used on each COM method exit - catch
------------------------------------------------------------------------*/

#define LEAVE_METHOD()								\
	}												\
	catch (_com_error& e)							\
	{												\
		SetErrorInfo(0, e.ErrorInfo());				\
		return e.Error();							\
	}												\
	catch (...)										\
	{												\
		return E_FAIL;								\
	}

/**----------------------------------------------------------------------
	Handy macro to return from method on invalid (NULL) parameter.
	that is not allowed to be NULL.
------------------------------------------------------------------------*/

#define countof(x)	(sizeof(x)/sizeof(x[0]))

#define ASSERT_RETURN(e)				ATLASSERT(e); if (!(e)) return;
#define ASSERT_CONTINUE(e)				ATLASSERT(e); if (!(e)) continue;
#define ASSERT_RETURN_VALUE(e,val)		ATLASSERT(e); if (!(e)) return (val);

#define DEFAULT_SOURCE	L"DockingShapeSheet Addin"

CString FormatString(LPCWSTR  format, ...);
void	ThrowComError (HRESULT hr, LPCWSTR error_message, LPCWSTR error_source = DEFAULT_SOURCE);
CString	FormatErrorMessage (_com_error& e);

CString LoadTextFromModule (HMODULE hModule, UINT resource_id);
BOOL LoadResourceFromModule(HMODULE hModule, LPCWSTR name, LPCWSTR type, LPBYTE* content, DWORD* content_length);
