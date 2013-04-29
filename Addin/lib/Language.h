
#pragma once

#define DEFAULT_LANGUAGE 1031

int GetAppLanguage(Visio::IVApplicationPtr app);

struct LanguageLock
{
	int old_lcid;
	int old_langid;

	LanguageLock(int app_language);
	~LanguageLock();
};

