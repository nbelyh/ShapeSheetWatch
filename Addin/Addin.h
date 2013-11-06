
#pragma once

#include "ViewSettings.h"

class CVisioFrameWnd;

class CAddinApp : public CWinApp
{
public:
	void OnCommand(UINT id);

	IVApplicationPtr GetVisioApp();
	void SetVisioApp(IVApplicationPtr app);

	virtual BOOL InitInstance();
	virtual int ExitInstance();

	ViewSettings* GetViewSettings() const;

private:
	CSimpleMap<HWND, CVisioFrameWnd*> m_shown_windows;

	mutable ViewSettings m_view_settings;

	IVApplicationPtr m_app;
};

extern CAddinApp theApp;
