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

	return TRUE;
}

int CAddinApp::ExitInstance() 
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

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

Office::IRibbonUIPtr CAddinApp::GetRibbon()
{
	return m_ribbon;
}

void CAddinApp::SetRibbon(Office::IRibbonUIPtr ribbon)
{
	m_ribbon = ribbon;
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

	if (m_ribbon)
		m_ribbon->Invalidate();
}

CAddinApp theApp;
