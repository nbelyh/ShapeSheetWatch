// AddIn.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "Addin.h"

#include "AddIn_i.h"
#include "AddIn_i.c"

#include "lib/Visio.h"
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
END_OBJECT_MAP()

BOOL CAddinApp::InitInstance()
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// Initialize COM stuff
	if (FAILED(_Module.Init(ObjectMap, AfxGetInstanceHandle(), &LIBID_AddinLib)))
		return FALSE;

	m_view_settings.Load();

	return TRUE;
}

int CAddinApp::ExitInstance() 
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	m_view_settings.Save();

	_Module.Term();

	return CWinApp::ExitInstance();
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

			HWND hwnd = GetVisioWindowHandle(window);

			CVisioFrameWnd* wnd = new CVisioFrameWnd();
			wnd->Create(window);
		}
	}
}

IVApplicationPtr CAddinApp::GetVisioApp()
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
