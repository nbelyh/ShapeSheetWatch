
#pragma once

#include "ViewSettings.h"

enum UpdateOption
{
	UpdateOption_UseKey = 1,
	UpdateOption_Hilight = 2,
	UpdateOption_Filter = 4,
	UpdateOption_Columns = 8,
	UpdateOption_Pin = 16,
};

struct IView
{
	virtual void Update(int hint) = 0;
};

struct VisioIdleTask
{
	virtual bool Execute() = 0;
	virtual bool Equals(VisioIdleTask* otehr) const = 0;
};

typedef Ptr<VisioIdleTask> VisioIdleTaskPtr;

class CVisioFrameWnd;

struct IAddinUI
{
	virtual void Update() = 0;
};

class CAddinApp 
	: public CWinApp
{
public:
	void OnCommand(UINT id);

	IVApplicationPtr GetVisioApp() const;
	void SetVisioApp(IVApplicationPtr app);

	IAddinUI* GetUI();
	void SetUI(IAddinUI* ui);

	virtual BOOL InitInstance();
	virtual int ExitInstance();

	ViewSettings* GetViewSettings() const;

	void AddView(IView* view);
	void DelView(IView* view);

	void UpdateViews(int hint = 0, bool idle = true);

	void AddVisioIdleTask(VisioIdleTaskPtr task);

	void ProcessIdleTasks();

	CVisioFrameWnd* GetWindowShapeSheet(IVWindowPtr window) const;
	void RegisterWindow(IVWindowPtr hwnd, CVisioFrameWnd* window);
	void UpdateVisioUI();

	virtual bool IsCommandEnabled(UINT id) const;
	virtual bool IsCommandVisible(UINT id) const;
	virtual bool IsCommandChecked(UINT id) const;

	bool IsShapeSheetWatchWindowShown(IVWindowPtr window) const;
	CVisioFrameWnd* ShowShapeSheetWatchWindow(IVWindowPtr window, bool show);

private:
	IVWindowPtr GetValidActiveWindow(VisWinTypes expected_type) const;

	CSimpleMap<HWND, CVisioFrameWnd*> m_shown_windows;

	typedef std::set<IView*> Views;
	Views m_views;

	CSimpleArray<VisioIdleTaskPtr> m_idle_tasks;

	mutable ViewSettings m_view_settings;

	IVApplicationPtr m_app;
	IAddinUI* m_ui;
};

extern CAddinApp theApp;

int GetVisioVersion();
