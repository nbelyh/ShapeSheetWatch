
#pragma once

#include "ViewSettings.h"

enum UpdateHint
{
	UpdateHint_Columns,

	UpdateHint_Count
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

class CAddinApp : public CWinApp
{
public:
	void OnCommand(UINT id);

	IVApplicationPtr GetVisioApp();
	void SetVisioApp(IVApplicationPtr app);

	virtual BOOL InitInstance();
	virtual int ExitInstance();

	ViewSettings* GetViewSettings() const;

	void AddView(IView* view);
	void DelView(IView* view);

	void UpdateViews(int hint = 0);

	void AddVisioIdleTask(VisioIdleTask* task);
	void ProcessIdleTasks();
private:
	typedef std::set<IView*> Views;
	Views m_views;

	CSimpleArray<VisioIdleTask*> m_idle_tasks;

	mutable ViewSettings m_view_settings;

	IVApplicationPtr m_app;
};

extern CAddinApp theApp;

int GetVisioVersion();
