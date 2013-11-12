
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

#define ASSERT_RETURN(e)				ATLASSERT(e); if (!(e)) return;
#define ASSERT_CONTINUE(e)				ATLASSERT(e); if (!(e)) continue;
#define ASSERT_RETURN_VALUE(e,val)		ATLASSERT(e); if (!(e)) return (val);

#define DEFAULT_SOURCE	L"ShapeSheetWatch Addin"

CString FormatString(LPCWSTR  format, ...);
void	ThrowComError (HRESULT hr, LPCWSTR error_message, LPCWSTR error_source = DEFAULT_SOURCE);
CString	FormatErrorMessage (_com_error& e);

CString LoadTextFromModule (HMODULE hModule, UINT resource_id);
BOOL LoadResourceFromModule(HMODULE hModule, LPCWSTR name, LPCWSTR type, LPBYTE* content, DWORD* content_length);

bstr_t QuoteString(LPCWSTR src);
CString UnquoteString(LPCWSTR src);
bool StringIsLike(LPCWSTR mask, LPCWSTR s);

CString GetListSep(bool add_space);

template <typename T>
struct GetSelf
{
	const T& operator() (const T& t) { return t; }
};

template <class T, class Fn>
CString JoinList(const T& src, LPCWSTR sep, Fn fn)
{
	CString result;

	T::const_iterator it = src.begin();
	while (it != src.end())
	{
		CString value = fn(*it);
		result += value;
		++it;

		if (it == src.end())
			break;

		if (sep != NULL)
			result += sep;
		else
			result += GetListSep(true);
	}

	return result;
}

template <class T>
CString JoinList(const T& src, LPCWSTR sep)
{
	return JoinList(src, sep, GetSelf<CString>());
}

template <typename T, typename Fn>
void SplitList(CString src, LPCWSTR sep, Fn fn, T& result)
{
	int fragment_begin = 0;
	while (fragment_begin < src.GetLength())
	{
		int fragment_end = src.Find(sep, fragment_begin);
		if (fragment_end < 0)
			fragment_end = src.GetLength();

		CString tok = 
			src.Mid(fragment_begin, fragment_end - fragment_begin);

		result.insert(result.end(), fn(tok));

		fragment_begin = fragment_end + int(wcslen(sep));
	}
}

template <typename T>
void SplitList(CString src, LPCWSTR sep, T& result)
{
	SplitList(src, sep, GetSelf<CString>(), result);
}
