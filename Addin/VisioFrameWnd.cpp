
/*******************************************************************************

	@file Implementation of a Visio frame window.

*******************************************************************************/

#include "stdafx.h"
#include "VisioFrameWnd.h"
#include "lib/Visio.h"
#include "Addin.h"

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

void CVisioFrameWnd::Create(IVWindowPtr window)
{
	HWND hwnd_parent = 
		::GetParent(GetVisioWindowHandle(window));

	CWnd* parent_window = 
		CWnd::FromHandle(hwnd_parent);

	CRect parent_rect;
	parent_window->GetClientRect(&parent_rect);

	// Construct Visio window. Make this window size a half of Visio's size
	IVWindowPtr this_window = window->GetWindows()->Add(
		bstr_t(L"Shape Sheet Watch"), 
		static_cast<long>(visWSVisible | visWSAnchorRight | visWSAnchorTop), 
		static_cast<long>(visAnchorBarAddon), 
		static_cast<long>(parent_rect.Width()), 
		static_cast<long>(parent_rect.Height()), 
		static_cast<long>(parent_rect.Width() / 4), 
		static_cast<long>(parent_rect.Height() / 2), 
		vtMissing, vtMissing, vtMissing);

	HWND client = GetVisioWindowHandle(this_window);
	SubclassWindow(client);

 	CRect rect;
 	GetClientRect(rect);
	m_html.Create(rect, this, 1, WS_CHILD|WS_VISIBLE, window);

	m_html.LoadHtml(
		L"<h2>HELLO</h2>"
		L"<widget style='width:100%;height:100%' type='shapesheet'></widget>"
		);
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

void CVisioFrameWnd::OnSize(UINT nType, int cx, int cy)
{
	m_html.MoveWindow(0, 0, cx, cy);
}

/**-----------------------------------------------------------------------------
	All background is filled with child frame - nothing to erase.
------------------------------------------------------------------------------*/

BOOL CVisioFrameWnd::OnEraseBkgnd(CDC* pDC)
{
	return TRUE;
}
