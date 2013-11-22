
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

private:
	typedef std::set<IView*> Views;
	Views m_views;

	mutable ViewSettings m_view_settings;

	IVApplicationPtr m_app;
};

extern CAddinApp theApp;

int GetVisioVersion();
