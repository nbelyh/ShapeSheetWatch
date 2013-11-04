
#pragma once

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

		fragment_begin = fragment_end + wcslen(sep);
	}
}

template <typename T>
void SplitList(CString src, LPCWSTR sep, T& result)
{
	SplitList(src, sep, GetSelf<CString>(), result);
}
