
#pragma once

#include "ViewSettings.h"

class CVisioFrameWnd;

class CAddinApp : public CWinApp
{
public:
	void OnCommand(UINT id);

	IVApplicationPtr GetVisioApp();
	void SetVisioApp(IVApplicationPtr app);

	Office::IRibbonUIPtr GetRibbon();
	void SetRibbon(Office::IRibbonUIPtr ribbon);

	virtual BOOL InitInstance();
	virtual int ExitInstance();

	ViewSettings* GetViewSettings() const;

	CVisioFrameWnd* GetWindowShapeSheet(HWND hwnd) const;
	void RegisterWindow(HWND hwnd, CVisioFrameWnd* window);

private:
	CSimpleMap<HWND, CVisioFrameWnd*> m_shown_windows;

	mutable ViewSettings m_view_settings;

	IVApplicationPtr m_app;
	Office::IRibbonUIPtr m_ribbon;
};

extern CAddinApp theApp;
