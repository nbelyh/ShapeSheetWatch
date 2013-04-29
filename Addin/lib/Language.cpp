
#include "stdafx.h"
#include "Language.h"

LanguageLock::LanguageLock( int app_language )
{
	HMODULE hKernel32 = GetModuleHandle(L"Kernel32.dll");

	typedef LANGID (WINAPI *FP_SetThreadUILanguage)(LANGID LangId);
	FP_SetThreadUILanguage pSetThreadUILanguage = (FP_SetThreadUILanguage)GetProcAddress(hKernel32, "SetThreadUILanguage");

	typedef LANGID (WINAPI *FP_GetThreadUILanguage)();
	FP_GetThreadUILanguage pGetThreadUILanguage = (FP_GetThreadUILanguage)GetProcAddress(hKernel32, "GetThreadUILanguage");

	old_lcid = GetThreadLocale();
	SetThreadLocale(app_language);

	old_langid = 0;
	if (pSetThreadUILanguage && pGetThreadUILanguage)
	{
		old_langid = pGetThreadUILanguage();
		pSetThreadUILanguage(app_language);
	}
}

LanguageLock::~LanguageLock()
{
	SetThreadLocale(old_lcid);

	HMODULE hKernel32 = GetModuleHandle(L"Kernel32.dll");

	typedef LANGID (WINAPI *FP_SetThreadUILanguage)(LANGID LangId);
	FP_SetThreadUILanguage pSetThreadUILanguage = (FP_SetThreadUILanguage)GetProcAddress(hKernel32, "SetThreadUILanguage");

	typedef LANGID (WINAPI *FP_GetThreadUILanguage)();
	FP_GetThreadUILanguage pGetThreadUILanguage = (FP_GetThreadUILanguage)GetProcAddress(hKernel32, "GetThreadUILanguage");

	if (pSetThreadUILanguage && pGetThreadUILanguage)
		pSetThreadUILanguage(old_langid);
}

int GetAppLanguage(Visio::IVApplicationPtr app)
{
	long app_language = 0;
	if (FAILED(app->get_Language(&app_language)))
		return DEFAULT_LANGUAGE;

	CComVariant v_disp_language_settings;
	CComDispatchDriver disp = app;
	if (SUCCEEDED(disp.GetPropertyByName(L"LanguageSettings", &v_disp_language_settings)))
	{
		Office::LanguageSettingsPtr language_settings;
		if (SUCCEEDED(V_DISPATCH(&v_disp_language_settings)->QueryInterface(__uuidof(Office::LanguageSettings), (void**)&language_settings)))
		{
			int ui_language = app_language;
			if (SUCCEEDED(language_settings->get_LanguageID(Office::msoLanguageIDUI, &ui_language)))
				app_language = ui_language;
		}
	}

	switch (app_language)
	{
	case 1033:
	case 1031:
	case 1049:
		return app_language;
	}

	return DEFAULT_LANGUAGE;
}
