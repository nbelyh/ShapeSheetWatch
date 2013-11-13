
// HTMLayoutCtrl.cpp : implementation file
//

#include "stdafx.h"

#include "lib/Utils.h"
#include "HTMLayoutCtrl.h"

#pragma warning( disable : 4996 )
#pragma warning( disable : 4267 )

#include <htmlayout.h>
#include <behaviors/behavior_popup.cpp>

#pragma warning( default : 4267 )
#pragma warning( default : 4996 )

/**-----------------------------------------------------------------------------
	Message to send to parent frame on hyperlink clicks
------------------------------------------------------------------------------*/

const UINT MSG_HTMLAYOUT_HYPERLINK = RegisterWindowMessage(L"P4B_MSG_HTMLAYOUT_HYPERLINK");
const UINT MSG_HTMLAYOUT_BUTTON = RegisterWindowMessage(L"P4B_MSG_HTMLAYOUT_BUTTON");

// CHTMLayoutCtrl

struct CHTMLayoutCtrl::Impl
{
	IHTMLayoutControlManager* m_control_manager;

	Impl(IHTMLayoutControlManager* control_manager)
		: m_control_manager(control_manager)
	{
	}

	void Attach(HWND hwnd)
	{
		ASSERT_RETURN(::IsWindow(hwnd));

		m_hwnd = hwnd;

		::HTMLayoutSetCallback(m_hwnd, callback, this);
		::HTMLayoutWindowAttachEventHandler(m_hwnd, &Impl::element_proc, this, HANDLE_BEHAVIOR_EVENT);
	}

	void Detach()
	{
		::HTMLayoutWindowDetachEventHandler(m_hwnd, &Impl::element_proc, this);

		m_hwnd = NULL;
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
		if (m_control_manager)
		{
			CString type = GetElemAttribute(lpCC->helement, "type");

			if (!type.IsEmpty())
			{
				CWnd* wnd = m_control_manager->CreateControl(type);

				if (wnd)
				{
					lpCC->outControlHwnd = wnd->GetSafeHwnd();
					return lpCC->outControlHwnd;
				}
			}
		}

		return HWND_TRY_DEFAULT;
	}

	LRESULT OnDestroyControl(LPNMHL_DESTROY_CONTROL lpDC)
	{
		if (m_control_manager)
		{
			 CString type = GetElemAttribute(lpDC->helement, "type");

			 CWnd* wnd = CWnd::FromHandle(lpDC->inoutControlHwnd);
			 
			 if (m_control_manager->DestroyControl(type, wnd))
			 {
				 lpDC->inoutControlHwnd = NULL;
				 return 0;
			 }
		}

		return 0;
	}

	LRESULT OnAttachBehavior(LPNMHL_ATTACH_BEHAVIOR lpab)
	{
		htmlayout::event_handler *pb = htmlayout::behavior::find(lpab->behaviorName, lpab->element);
		if (pb)
		{
			lpab->elementTag  = pb;
			lpab->elementProc = htmlayout::behavior::element_proc;
			lpab->elementEvents = pb->subscribed_to;
		}
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
		case HLN_ATTACH_BEHAVIOR:   return OnAttachBehavior((LPNMHL_ATTACH_BEHAVIOR)lParam);
		}

		return OnHtmlGenericNotifications(uMsg,wParam,lParam);
	}

	static BOOL CALLBACK FindElementCallback(HELEMENT he, LPVOID param)
	{
		*reinterpret_cast<HELEMENT*>(param) = he;
		return TRUE;
	}

	CString GetElemType(HELEMENT elem)
	{
		LPCSTR val = NULL;
		HTMLayoutGetElementType(elem, &val);
		return val ? CString(val) : L"";
	}

	CString GetElemAttribute(HELEMENT elem, const char* attribute)
	{
		LPCWSTR val = NULL;
		HTMLayoutGetAttributeByName(elem, attribute, &val);
		return val ? CString(val) : L"";
	}

	void SetElemAttribute(HELEMENT elem, const char* attribute, LPCWSTR val)
	{
		HTMLayoutSetAttributeByName(elem, attribute, val);
	}

	HELEMENT GetElemById(const char* id) const
	{
		HELEMENT h_root = 0;
		HTMLayoutGetRootElement(m_hwnd, &h_root);

		char selector[256] = "";
		sprintf_s(selector, 256, "#%s", id);

		HELEMENT h_found = 0;
		HTMLayoutSelectElements(h_root, selector, &FindElementCallback, &h_found);

		return h_found;
	}

	void SetElemEnabled(HELEMENT elem, const char* id, bool enabled)
	{
		if (enabled)
			HTMLayoutSetElementState(elem, 0, STATE_DISABLED, TRUE);
		else
			HTMLayoutSetElementState(elem, STATE_DISABLED, 0, TRUE);
	}

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

BOOL CHTMLayoutCtrl::Create(const RECT& rect, CWnd* pParentWnd, UINT nID, DWORD dwStyle)
{
	if (!CWnd::Create(::HTMLayoutClassNameT(), NULL, dwStyle, rect, pParentWnd, nID))
		return FALSE;

	m_impl->Attach(m_hWnd);
	return TRUE;
}

/**-----------------------------------------------------------------------------
	Constructor
------------------------------------------------------------------------------*/

CHTMLayoutCtrl::CHTMLayoutCtrl(IHTMLayoutControlManager* manager)
	: m_impl(new Impl(manager))
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

void CHTMLayoutCtrl::SetElementText(const char* id, LPCWSTR text)
{
	HELEMENT h_found = m_impl->GetElemById(id);
	HTMLayoutSetElementInnerText16(h_found, text, (UINT)lstrlenW(text));
}

CString CHTMLayoutCtrl::GetElementAttribute (const char* id, const char* attribute)
{
	HELEMENT h_found = m_impl->GetElemById(id);
	return m_impl->GetElemAttribute(h_found, attribute);
}

void CHTMLayoutCtrl::SetElementAttribute (const char* id, const char* attribute, LPCWSTR text)
{
	HELEMENT h_found = m_impl->GetElemById(id);
	m_impl->SetElemAttribute(h_found, attribute, text);
}

void CHTMLayoutCtrl::SetElementEnabled(const char* id, bool enabled)
{
	HELEMENT h_found = m_impl->GetElemById(id);
	m_impl->SetElemEnabled(h_found, id, enabled);
}

IMPLEMENT_DYNAMIC(CHTMLayoutCtrl, CWnd)
