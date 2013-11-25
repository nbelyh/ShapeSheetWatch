
/*******************************************************************************

	@file Implementation of a Visio frame window.

*******************************************************************************/

#include "stdafx.h"
#include "VisioFrameWnd.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
#include "ShapeSheetGridCtrl.h"
#include "Addin.h"

#define WINDOW_CAPTION		L"Shape Sheet Watch"

/**-----------------------------------------------------------------------------
	Message map
------------------------------------------------------------------------------*/

BEGIN_MESSAGE_MAP(CVisioFrameWnd, CWnd)
	ON_WM_CREATE()
	ON_WM_ERASEBKGND()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
END_MESSAGE_MAP()

/**------------------------------------------------------------------------------
	Creates and initializes a new Visio (top-level) window, and subclasses the 
	"client area" of that parent Visio window (this "client area" is actually 
	a special window provided by visio for sub-classing).
-------------------------------------------------------------------------------*/

CString GetHtmlFilePath(LPCWSTR file_name)
{
	WCHAR path[MAX_PATH] = L"";
	DWORD path_len = MAX_PATH;

	if (0 == SHRegGetUSValue(L"Software\\Microsoft\\Visio\\Addins\\ShapeSheetWatchAddin.VisioConnect", L"InstallPath", 
		NULL, path, &path_len, FALSE, NULL, 0))
	{
		PathAddBackslash(path);
		PathAppend(path, L"html");
		PathAddBackslash(path);
		PathAppend(path, file_name);
	}

	return path;
}

void CVisioFrameWnd::Create(IVWindowPtr window)
{
	m_window = window;

	long l = 0, t = 0, w = 400, h = 300;
	theApp.GetViewSettings()->GetWindowRect(l, t, w, h);

	long state = theApp.GetViewSettings()->GetWindowState();

	// Construct Visio window. Make this window size a half of Visio's size
	m_this_window = window->GetWindows()->Add(
		bstr_t(WINDOW_CAPTION), 
		state, static_cast<long>(visAnchorBarAddon), 
		l, t, w, h,
		vtMissing, vtMissing, vtMissing);

	HWND client = GetVisioWindowHandle(m_this_window);
	SubclassWindow(client);

	CRect client_rect;
 	GetClientRect(client_rect);
	m_html.Create(client_rect, this, 1, WS_CHILD|WS_VISIBLE|WS_CLIPCHILDREN);

	CString index_html = GetHtmlFilePath(L"test.htm");
	m_html.LoadHtmlFile(index_html);

	SetChecks();
}

void CVisioFrameWnd::Destroy()
{
	GetParent()->SendMessage(WM_CLOSE);
}

/**-----------------------------------------------------------------------------
	Called when window is about destroying - save window position.
------------------------------------------------------------------------------*/

void CVisioFrameWnd::PostNcDestroy()
{
	delete this;
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

BOOL CVisioFrameWnd::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	return TRUE;
}

/**-----------------------------------------------------------------------------
	Called when window size is changed.
------------------------------------------------------------------------------*/

struct UpdateWindowSize : VisioIdleTask
{
	UpdateWindowSize(IVWindowPtr this_window)
		: m_this_window(this_window)
	{
	}

	virtual bool Execute()
	{
		long l = 0, t = 0, w = 400, h = 300;
		if (SUCCEEDED(m_this_window->raw_GetWindowRect(&l, &t, &w, &h)))
			theApp.GetViewSettings()->SetWindowRect(l, t, w, h);

		long state = (visWSVisible | visWSAnchorTop | visWSAnchorRight);
		if (SUCCEEDED(m_this_window->get_WindowState(&state)))
			theApp.GetViewSettings()->SetWindowState(state);

		return true;
	}

	virtual bool Equals(VisioIdleTask* otehr) const
	{
		UpdateWindowSize* other_task = dynamic_cast<UpdateWindowSize*>(otehr);
		if (other_task)
			return m_this_window->ID == other_task->m_this_window->ID;
		else
			return false;
	}

	IVWindowPtr m_this_window;
};


void CVisioFrameWnd::OnSize(UINT nType, int cx, int cy)
{
	m_html.MoveWindow(0, 0, cx, cy);

	theApp.AddVisioIdleTask(new UpdateWindowSize(m_this_window));
}

/**-----------------------------------------------------------------------------
	All background is filled with child frame - nothing to erase.
------------------------------------------------------------------------------*/

BOOL CVisioFrameWnd::OnEraseBkgnd(CDC* pDC)
{
	return TRUE;
}

CVisioFrameWnd::CVisioFrameWnd()
	: m_html(this)
{
}

CWnd* CVisioFrameWnd::CreateControl(const CString& type)
{
	if (type != L"sheet")
		return NULL;

	CShapeSheetGridCtrl* pShapeSheetCtrl = new CShapeSheetGridCtrl();

	if (pShapeSheetCtrl->Create(&m_html, 1, m_window))
		return pShapeSheetCtrl;

	return NULL;
}

bool CVisioFrameWnd::DestroyControl(const CString& type, CWnd* wnd)
{
	if (type != L"sheet")
		return false;

	CShapeSheetGridCtrl* pShapeSheetCtrl = static_cast<CShapeSheetGridCtrl*>(wnd);

	pShapeSheetCtrl->Destroy();
	delete pShapeSheetCtrl;

	return true;
}

void CVisioFrameWnd::SetChecks()
{
	for (int i = 0; i < Column_Count; ++i)
	{
		bool visible = theApp.GetViewSettings()->IsColumnVisible(i);
		m_html.SetElementChecked(GetColumnDbName(i), visible);
	}
}

bool CVisioFrameWnd::OnCheckButton(const CString& id, bool visible)
{
	int col = GetColumnFromDbName(id);

	if (col >= 0)
	{
		theApp.GetViewSettings()->SetColumnVisible(col, visible);
		theApp.UpdateViews(UpdateHint_Columns);
	}

	return true;
}
