// AddIn.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "Addin.h"

#include "AddIn_i.h"
#include "AddIn_i.c"

#include "lib/Visio.h"
#include "VisioConnect.h"
#include "VisioFrameWnd.h"

CComModule _Module;

// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
	return _Module.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _Module.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
	// registers object, typelib and all interfaces in typelib
	HRESULT hr = _Module.DllRegisterServer();
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _Module.DllUnregisterServer();
	return hr;
}

STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	// MSVC will call "regsvr32 /i:user" if "per-user registration" is set as a
	// linker option - so handle that here (its also handle for anyone else to
	// be able to manually install just for themselves.)
	static const wchar_t szUserSwitch[] = L"user";
	if (pszCmdLine != NULL)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			AtlSetPerUserRegistration(true);
			// But ATL still barfs if you try and register a COM category, so
			// just arrange to not do that.
			_AtlComModule.m_ppAutoObjMapFirst = _AtlComModule.m_ppAutoObjMapLast;
		}
	}
	if (bInstall)
	{
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}
	return hr;
}

BEGIN_OBJECT_MAP(ObjectMap)
	OBJECT_ENTRY(__uuidof(VisioConnect), CVisioConnect)
END_OBJECT_MAP()

BOOL CAddinApp::InitInstance()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// Initialize COM stuff
	if (FAILED(_Module.Init(ObjectMap, AfxGetInstanceHandle(), &LIBID_AddinLib)))
		return FALSE;

	return TRUE;
}

int CAddinApp::ExitInstance() 
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	_Module.Term();

	return CWinApp::ExitInstance();
}

bool CAddinApp::IsShapeSheetWatchWindowShown(IVWindowPtr window) const
{
	return GetWindowShapeSheet(GetVisioWindowHandle(window)) != NULL;
}

void CAddinApp::ShowShapeSheetWatchWindow(IVWindowPtr window, bool show)
{
	bool shown = IsShapeSheetWatchWindowShown(window);

	if (show == shown)
		return;

	HWND hwnd = GetVisioWindowHandle(window);

	CVisioFrameWnd* wnd = GetWindowShapeSheet(hwnd);
	if (wnd)
	{
		wnd->Destroy();
		RegisterWindow(hwnd, NULL);
	}
	else
	{
		wnd = new CVisioFrameWnd();
		wnd->Create(window);
		RegisterWindow(hwnd, wnd);
	}
}

void CAddinApp::OnCommand(UINT id)
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		{
			IVWindowPtr window = m_app->GetActiveWindow();

			if (window == NULL)
				return;

			ShowShapeSheetWatchWindow(window, !IsShapeSheetWatchWindowShown(window));
		}
	}
}

IVApplicationPtr CAddinApp::GetVisioApp() const
{
	return m_app;
}

void CAddinApp::SetVisioApp( IVApplicationPtr app )
{
	m_app = app;
}

ViewSettings* CAddinApp::GetViewSettings() const
{
	return &m_view_settings;
}

void CAddinApp::AddView( IView* view )
{
	m_views.insert(view);
}

void CAddinApp::DelView( IView* view )
{
	m_views.erase(view);
}

void CAddinApp::UpdateViews(int hint)
{
	for (Views::const_iterator it = m_views.begin(); it != m_views.end(); ++it)
		(*it)->Update(hint);
}

CAddinApp theApp;

int GetVisioVersion()
{
	static int result = -1;

	if (result == -1)
		result = StrToInt(theApp.GetVisioApp()->GetVersion());

	return result;
}

void CAddinApp::AddVisioIdleTask(VisioIdleTask* new_task)
{
	long idx = 0;
	while (idx < m_idle_tasks.GetSize())
	{
		VisioIdleTask* task = m_idle_tasks[idx];

		if (new_task->Equals(task))
		{
			delete new_task;
			return;
		}
		else
		{
			++idx;
		}
	}

	m_idle_tasks.Add(new_task);
}

void CAddinApp::ProcessIdleTasks()
{
	long idx = 0;
	while (idx < m_idle_tasks.GetSize())
	{
		VisioIdleTask* task = m_idle_tasks[idx];

		if (task->Execute())
		{
			m_idle_tasks.RemoveAt(idx);
			delete task;
		}
		else
		{
			++idx;
		} 
	}
}

CVisioFrameWnd* CAddinApp::GetWindowShapeSheet(HWND hwnd) const
{
	int idx = m_shown_windows.FindKey(hwnd);

	if (idx < 0)
		return NULL;
	return
		m_shown_windows.GetValueAt(idx);
}

void CAddinApp::RegisterWindow(HWND hwnd, CVisioFrameWnd* window)
{
	if (window)
		m_shown_windows.Add(hwnd, window);
	else
		m_shown_windows.Remove(hwnd);

	UpdateVisioUI();
}

IAddinUI* CAddinApp::GetUI()
{
	return m_ui;
}

void CAddinApp::SetUI(IAddinUI* ui)
{
	m_ui = ui;
}

struct UpdateUITask : VisioIdleTask
{
	virtual bool Execute()
	{
		theApp.GetUI()->Update();
		return true;
	}

	virtual bool Equals(VisioIdleTask* otehr) const
	{
		return dynamic_cast<UpdateUITask*>(otehr) != NULL;
	}
};

void CAddinApp::UpdateVisioUI()
{
	AddVisioIdleTask(new UpdateUITask());
}

bool CAddinApp::IsCommandEnabled(UINT id) const
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		{
			IVApplicationPtr app = GetVisioApp();

			IVDocumentPtr doc;
			if (FAILED(app->get_ActiveDocument(&doc)) || doc == NULL)
				return false;

			VisDocumentTypes doc_type = visDocTypeInval;
			if (FAILED(doc->get_Type(&doc_type)) || doc_type == visDocTypeInval)
				return false;

			return true;
		}

	default:
		return true;
	}
}

bool CAddinApp::IsCommandVisible(UINT id) const
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		{
			return true;
		}

	default:
		return true;
	}
}

bool CAddinApp::IsCommandChecked(UINT id) const
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		{
			IVApplicationPtr app = GetVisioApp();

			IVWindowPtr window;
			if (FAILED(app->get_ActiveWindow(&window)) || window == NULL)
				return VARIANT_FALSE;

			return IsShapeSheetWatchWindowShown(window);
		}

	default:
		return false;
	}
}
