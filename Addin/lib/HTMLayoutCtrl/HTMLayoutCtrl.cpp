
// HTMLayoutCtrl.cpp : implementation file
//

#include "stdafx.h"

#include "lib/Utils.h"
#include "ShapeSheetGridCtrl.h"
#include "HTMLayoutCtrl.h"

/**-----------------------------------------------------------------------------
	Message to send to parent frame on hyperlink clicks
------------------------------------------------------------------------------*/

const UINT MSG_HTMLAYOUT_HYPERLINK = RegisterWindowMessage(L"P4B_MSG_HTMLAYOUT_HYPERLINK");
const UINT MSG_HTMLAYOUT_BUTTON = RegisterWindowMessage(L"P4B_MSG_HTMLAYOUT_BUTTON");

// CHTMLayoutCtrl

struct CHTMLayoutCtrl::Impl
{
	void Attach(HWND hwnd, IVWindowPtr window)
	{
		ASSERT_RETURN(::IsWindow(hwnd));

		m_hwnd = hwnd;
		m_window = window;

		::HTMLayoutSetCallback(m_hwnd, callback, this);
		::HTMLayoutWindowAttachEventHandler(m_hwnd, &Impl::element_proc, this, HANDLE_BEHAVIOR_EVENT);
	}

	void Detach()
	{
		::HTMLayoutWindowDetachEventHandler(m_hwnd, &Impl::element_proc, this);

		m_hwnd = NULL;
		m_window = NULL;
	}

	/**-----------------------------------------------------------------------------
		HTML behavior events handler
	------------------------------------------------------------------------------*/

	CWnd* FindOwnerWindow()
	{
		CWnd* frame = 
			CWnd::FromHandle(m_hwnd)->GetParentFrame();

		if (frame == NULL)
			frame = CWnd::FromHandle(m_hwnd)->GetParent();

		return frame;
	}

	BOOL handle_event (HELEMENT he, BEHAVIOR_EVENT_PARAMS* params) 
	{
		if (!IsWindow(m_hwnd))
			return FALSE;

		if (params->cmd == BUTTON_CLICK)
		{
			LPCWSTR id = L"";
			::HTMLayoutGetAttributeByName(he, "id", &id);

			CWnd* frame = FindOwnerWindow();
			if (frame != NULL)
				frame->SendMessage(MSG_HTMLAYOUT_BUTTON, reinterpret_cast<LPARAM>(id), NULL);
		}
		if (params->cmd == HYPERLINK_CLICK)
		{
			LPCWSTR href = L"";
			::HTMLayoutGetAttributeByName(he, "href", &href);

			LPCWSTR id = L"";
			::HTMLayoutGetAttributeByName(he, "id", &id);

			CWnd* frame = FindOwnerWindow();
			if (frame != NULL)
				frame->SendMessage(MSG_HTMLAYOUT_HYPERLINK, reinterpret_cast<WPARAM>(id), reinterpret_cast<LPARAM>(href));
		}

		return FALSE;
	}

	/**-----------------------------------------------------------------------------
		HTML behavior callback
	------------------------------------------------------------------------------*/

	static BOOL CALLBACK element_proc(LPVOID tag, HELEMENT he, UINT evtg, LPVOID prms)
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState());

		Impl* p_this = reinterpret_cast<Impl*>(tag);
		return p_this->handle_event(he, (BEHAVIOR_EVENT_PARAMS*) prms);
	}

	/**-----------------------------------------------------------------------------
		non-HTMLayout notifications	
	------------------------------------------------------------------------------*/

	LRESULT OnHtmlGenericNotifications(UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		// Pass it to the parent window if any
		HWND hWndParent = ::GetParent(m_hwnd);
		if (!hWndParent) return 0;

		return ::SendMessage(hWndParent, uMsg, wParam, lParam);
	}

	/**-----------------------------------------------------------------------------
		HTMLayout callback
	------------------------------------------------------------------------------*/

	static LRESULT CALLBACK callback(UINT uMsg, WPARAM wParam, LPARAM lParam, LPVOID vParam)
	{
		ASSERT(vParam);
		Impl* pThis = (Impl*)vParam;
		return pThis->OnHtmlNotify(uMsg, wParam, lParam);
	}

	/**-----------------------------------------------------------------------------
		
	------------------------------------------------------------------------------*/

	HWND OnCreateControl(LPNMHL_CREATE_CONTROL lpCC)
	{
		CShapeSheetGridCtrl* pShapeSheetCtrl = 
			new CShapeSheetGridCtrl();

		if (!pShapeSheetCtrl->Create(CWnd::FromHandle(m_hwnd), 1, m_window))
			return HWND_DISCARD_CREATION;

		lpCC->outControlHwnd = pShapeSheetCtrl->GetSafeHwnd();

		return lpCC->outControlHwnd;
	}

	LRESULT OnDestroyControl(LPNMHL_DESTROY_CONTROL lpDC)
	{
		CShapeSheetGridCtrl* pShapeSheetCtrl = 
			static_cast<CShapeSheetGridCtrl*>(CWnd::FromHandle(lpDC->inoutControlHwnd));

		pShapeSheetCtrl->Detach();
		delete pShapeSheetCtrl;

		lpDC->inoutControlHwnd = NULL;

		return 0;
	}

	LRESULT OnHtmlNotify(UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		ASSERT(uMsg == WM_NOTIFY);

		// Crack and call appropriate method
		switch(((NMHDR*)lParam)->code) 
		{
		case HLN_CREATE_CONTROL:    return reinterpret_cast<LRESULT>(OnCreateControl(reinterpret_cast<LPNMHL_CREATE_CONTROL>(lParam)));
		case HLN_CONTROL_CREATED:   return 0;
		case HLN_DESTROY_CONTROL:   return OnDestroyControl(reinterpret_cast<LPNMHL_DESTROY_CONTROL>(lParam));
		case HLN_LOAD_DATA:         return 0;
		case HLN_DATA_LOADED:       return 0;
		case HLN_DOCUMENT_COMPLETE: return 0;
		case HLN_ATTACH_BEHAVIOR:   return 0;
		}

		return OnHtmlGenericNotifications(uMsg,wParam,lParam);
	}

	IVWindowPtr m_window;
	HWND m_hwnd;
};

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

bool CHTMLayoutCtrl::LoadHtml(const BYTE * pb, DWORD nBytes)
{
	ASSERT_RETURN_VALUE(::IsWindow(m_hWnd), false);
	return ::HTMLayoutLoadHtml(m_hWnd, pb, nBytes) != 0;
}

bool CHTMLayoutCtrl::LoadHtml(LPCWSTR text)
{
	return LoadHtml(reinterpret_cast<const BYTE*>(text), 2*lstrlen(text));
}

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

bool CHTMLayoutCtrl::LoadHtmlEx(const BYTE * pb, DWORD nBytes, LPCWSTR base_path)
{
	ASSERT_RETURN_VALUE(::IsWindow(m_hWnd), false);

	return 0 != ::HTMLayoutLoadHtmlEx(m_hWnd, pb, nBytes, base_path);
}

bool CHTMLayoutCtrl::LoadHtmlEx(LPCWSTR text, LPCWSTR base_path)
{
	return LoadHtmlEx(reinterpret_cast<const BYTE*>(text), 2*lstrlen(text), base_path);
}

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

bool CHTMLayoutCtrl::LoadHtmlFile(LPCWSTR filename)
{
	ASSERT_RETURN_VALUE(::IsWindow(m_hWnd), false);
	return ::HTMLayoutLoadFile(m_hWnd, filename) != 0;
}

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

BOOL CHTMLayoutCtrl::PreTranslateMessage(MSG* pMsg )
{
	if (pMsg->message == WM_CHAR)
	{
		BOOL bHandled = FALSE;
		::HTMLayoutProcND(m_hWnd, pMsg->message, pMsg->wParam, pMsg->lParam, &bHandled);
		if(bHandled)
			return TRUE;
	}

	return CWnd::PreTranslateMessage(pMsg);
}

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

BOOL CHTMLayoutCtrl::Create(const RECT& rect, CWnd* pParentWnd, UINT nID, DWORD dwStyle, IVWindowPtr window)
{
	if (!CWnd::Create(::HTMLayoutClassNameT(), NULL, dwStyle, rect, pParentWnd, nID))
		return FALSE;

	m_impl->Attach(m_hWnd, window);
	return TRUE;
}

/**-----------------------------------------------------------------------------
	Constructor
------------------------------------------------------------------------------*/

CHTMLayoutCtrl::CHTMLayoutCtrl()
	: m_impl(new Impl())
{
}

/**-----------------------------------------------------------------------------
	Destructor
------------------------------------------------------------------------------*/

CHTMLayoutCtrl::~CHTMLayoutCtrl()
{
	delete m_impl;
}

UINT CHTMLayoutCtrl::GetMinHeight (UINT width) const
{
	return ::HTMLayoutGetMinHeight(m_hWnd, width);
}

UINT CHTMLayoutCtrl::GetMinWidth () const
{
	return ::HTMLayoutGetMinWidth(m_hWnd);
}

namespace {

	BOOL CALLBACK FindElementCallback(HELEMENT he, LPVOID param)
	{
		*reinterpret_cast<HELEMENT*>(param) = he;
		return TRUE;
	}

}
HELEMENT CHTMLayoutCtrl::GetElementById(const char* id) const
{
	HELEMENT h_root = 0;
	HTMLayoutGetRootElement(m_hWnd, &h_root);

	char selector[256] = "";
	sprintf_s(selector, 256, "#%s", id);

	HELEMENT h_found = 0;
	HTMLayoutSelectElements(h_root, selector, &FindElementCallback, &h_found);

	return h_found;
}

void CHTMLayoutCtrl::SetElementText(const char* id, LPCWSTR text)
{
	HELEMENT h_found = GetElementById(id);
	HTMLayoutSetElementInnerText16(h_found, text, (UINT)lstrlenW(text));
}

void CHTMLayoutCtrl::SetElementAttribute (const char* id, const char* attribute, LPCWSTR text)
{
	HELEMENT h_found = GetElementById(id);
	HTMLayoutSetAttributeByName(h_found, attribute, text);
}

void CHTMLayoutCtrl::SetElementEnabled(const char* id, bool enabled)
{
	HELEMENT h_found = GetElementById(id);

	if (enabled)
		HTMLayoutSetElementState(h_found, 0, STATE_DISABLED, TRUE);
	else
		HTMLayoutSetElementState(h_found, STATE_DISABLED, 0, TRUE);
}

IMPLEMENT_DYNAMIC(CHTMLayoutCtrl, CWnd)
