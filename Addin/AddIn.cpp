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
	return GetWindowShapeSheet(window) != NULL;
}

CVisioFrameWnd* CAddinApp::ShowShapeSheetWatchWindow(IVWindowPtr window, bool show)
{
	CVisioFrameWnd* wnd = GetWindowShapeSheet(window);

	if ((wnd != NULL) == show)
		return wnd;

	if (wnd)
	{
		wnd->Destroy();
		RegisterWindow(window, NULL);
		return NULL;
	}
	else
	{
		wnd = new CVisioFrameWnd();
		wnd->Create(window);
		RegisterWindow(window, wnd);
		return wnd;
	}
}

IVWindowPtr FindDocumentWindow(IVCellPtr cell)
{
	long doc_id = cell->GetDocument()->GetID();

	IVWindowsPtr windows = cell->Application->GetWindows();
	for (short i = 1; i <= windows->GetCount(); ++i)
	{
		IVWindowPtr window = windows->GetItem(i);

		if (window->GetDocument()->GetID() == doc_id)
			return window;
	}

	return NULL;
}

void CAddinApp::OnCommand(UINT id)
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		{
			IVWindowPtr window = GetValidActiveWindow(visDrawing);

			if (!window)
				return;

			ShowShapeSheetWatchWindow(window, !IsShapeSheetWatchWindowShown(window));
			break;
		}

	case ID_AddWatch:
		{
			IVWindowPtr sheet_window = GetValidActiveWindow(visSheet);

			if (!sheet_window)
				return;

			IVCellPtr selected_cell = sheet_window->GetSelectedCell();
			if (selected_cell)
			{
				IVWindowPtr window = FindDocumentWindow(selected_cell);

				if (window)
				{
					CVisioFrameWnd* wnd = 
						ShowShapeSheetWatchWindow(window, true);

					GetViewSettings()->AddCellMask(selected_cell->Name);
					UpdateViews(UpdateHint_Rows);
				}
			}
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

CVisioFrameWnd* CAddinApp::GetWindowShapeSheet(IVWindowPtr window) const
{
	HWND hwnd = GetVisioWindowHandle(window);

	int idx = m_shown_windows.FindKey(hwnd);

	if (idx < 0)
		return NULL;
	return
		m_shown_windows.GetValueAt(idx);
}

void CAddinApp::RegisterWindow(IVWindowPtr window, CVisioFrameWnd* sheet)
{
	HWND hwnd = GetVisioWindowHandle(window);

	if (sheet)
		m_shown_windows.Add(hwnd, sheet);
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

IVWindowPtr CAddinApp::GetValidActiveWindow(VisWinTypes expected_type) const
{
	IVApplicationPtr app = GetVisioApp();

	IVWindowPtr window;
	if (FAILED(app->get_ActiveWindow(&window)) || window == NULL)
		return NULL;

	short window_type = 0;
	if (FAILED(window->get_Type(&window_type)) || window_type != expected_type)
		return NULL;

	return window;
}

bool CAddinApp::IsCommandEnabled(UINT id) const
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		return GetValidActiveWindow(visDrawing) != NULL;

	case ID_AddWatch:
		return GetValidActiveWindow(visSheet) != NULL;

	default:
		return true;
	}
}

bool CAddinApp::IsCommandVisible(UINT id) const
{
	switch (id)
	{
	case ID_ShowSheetWindow:
		return true;

	case ID_AddWatch:
		return true;

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
			IVWindowPtr window = GetValidActiveWindow(visDrawing);

			if (!window)
				return VARIANT_FALSE;

			return IsShapeSheetWatchWindowShown(window);
		}

	case ID_AddWatch:
		{
			return false;
		}

	default:
		return false;
	}
}
