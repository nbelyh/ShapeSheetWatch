
// HTMLayoutCtrl.cpp : implementation file
//

#include "stdafx.h"

#include "HTMLayoutCtrl.h"
#include "ShapeSheetCtrl.h"

/**-----------------------------------------------------------------------------
	Message to send to parent frame on hyperlink clicks
------------------------------------------------------------------------------*/

const UINT MSG_HTMLAYOUT_HYPERLINK = RegisterWindowMessage(L"P4B_MSG_HTMLAYOUT_HYPERLINK");
const UINT MSG_HTMLAYOUT_BUTTON = RegisterWindowMessage(L"P4B_MSG_HTMLAYOUT_BUTTON");

// CHTMLayoutCtrl

struct CHTMLayoutCtrl::Impl
{
	CSimpleArray<CShapeSheetGrid*> m_grids;

	void Attach(HWND hwnd)
	{
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
		Attach your own behavior to the element
	------------------------------------------------------------------------------*/

	LRESULT OnDestroyControl(  LPNMHL_DESTROY_CONTROL lpdc)
	{
		CWnd* wnd = CWnd::FromHandlePermanent(lpdc->inoutControlHwnd);
		if (wnd)
		{
			CShapeSheetGrid* grid = DYNAMIC_DOWNCAST(CShapeSheetGrid, wnd);
			ASSERT_VALID(grid);

			m_grids.Remove(grid);
			delete grid;
		}

		return 0;
	}

	LRESULT OnCreateControl( LPNMHL_CREATE_CONTROL lpcc )
	{
		htmlayout::dom::element el(lpcc->helement);

		const wchar_t* type = el.get_attribute("type");
		if (type != NULL && !lstrcmp(type, L"grid"))
		{
			const wchar_t* name = el.get_attribute("name");
			const wchar_t* id = el.get_attribute("id");

			CShapeSheetGrid* grid = new CShapeSheetGrid(name, lpcc->helement);
			lpcc->outControlHwnd = grid->Create(lpcc->inHwndParent, StrToInt(id));

			grid->AutoSize();

			m_grids.Add(grid);
		}

		return 0;
	}

	LRESULT OnAttachBehavior( LPNMHL_ATTACH_BEHAVIOR lpab )
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

	LRESULT OnHtmlNotify(UINT uMsg, WPARAM wParam, LPARAM lParam)
	{
		ASSERT(uMsg == WM_NOTIFY);

		// Crack and call appropriate method
		switch(((NMHDR*)lParam)->code) 
		{
		case HLN_CREATE_CONTROL:    
			return OnCreateControl((LPNMHL_CREATE_CONTROL)lParam);

		case HLN_CONTROL_CREATED:   
			return 0;

		case HLN_DESTROY_CONTROL:   
			return OnDestroyControl((LPNMHL_DESTROY_CONTROL)lParam);

		case HLN_LOAD_DATA:         
			return 0;

		case HLN_DATA_LOADED:       
			return 0;

		case HLN_DOCUMENT_COMPLETE: 
			return 0;

		case HLN_ATTACH_BEHAVIOR:   
			return OnAttachBehavior((LPNMHL_ATTACH_BEHAVIOR)lParam);
		}

		return OnHtmlGenericNotifications(uMsg,wParam,lParam);
	}

	HWND m_hwnd;
};

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

bool CHTMLayoutCtrl::LoadHtml(const BYTE * pb, DWORD nBytes)
{
	return ::HTMLayoutLoadHtml(m_hWnd, pb, nBytes) != 0;
}

bool CHTMLayoutCtrl::LoadHtml(LPCWSTR text)
{
	return LoadHtml(reinterpret_cast<const BYTE*>(text), 2*lstrlen(text));
}

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

bool CHTMLayoutCtrl::LoadHtmlFile(LPCWSTR filename)
{
	return ::HTMLayoutLoadFile(m_hWnd, filename) != 0;
}

/**-----------------------------------------------------------------------------
	
------------------------------------------------------------------------------*/

BOOL CHTMLayoutCtrl::PreTranslateMessage(MSG* pMsg )
{
	if (pMsg->message == WM_CHAR || pMsg->message == WM_KEYDOWN || pMsg->message == WM_KEYUP)
	{
		BOOL bHandled = FALSE;
		::HTMLayoutProcND(m_hWnd, pMsg->message, pMsg->wParam, pMsg->lParam, &bHandled);
		if(bHandled) return TRUE;
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

using namespace htmlayout::dom;

void CHTMLayoutCtrl::SetElementText(const char* id, LPCWSTR text)
{
	element el_root = element::root_element(m_hWnd);
	element el = el_root.find_first(id);
	if (el)
		el.set_text(text);
}

void CHTMLayoutCtrl::SetElementAttribute (const char* id, const char* attribute, LPCWSTR text)
{
	element el_root = element::root_element(m_hWnd);
	element el = el_root.find_first(id);
	if (el)
		el.set_attribute(attribute, text);
}

void CHTMLayoutCtrl::SetElementEnabled (const char* id, bool set)
{
	element el_root = element::root_element(m_hWnd);
	element el = el_root.find_first(id);
	if (set)
		el.set_state(0, STATE_DISABLED);
	else
		el.set_state(STATE_DISABLED, 0);
}

void CHTMLayoutCtrl::SetElementValueLong(const char* id, long v)
{
	element el_root = element::root_element(m_hWnd);
	element el = el_root.find_first(id);
	el.set_value(json::value(v));
}

LRESULT CHTMLayoutCtrl::WindowProc(UINT message,WPARAM wParam,LPARAM lParam )
{
	BOOL handled = FALSE;
	BOOL scrolling =
		(message == WM_MOUSEWHEEL);

	if (scrolling)
	{
		SetRedraw(FALSE);
	}
	LRESULT lr = HTMLayoutProcND(m_hWnd, message, wParam, lParam, &handled);
	if (scrolling)
	{
		SetRedraw(TRUE);

		HELEMENT rt;
		HTMLayoutGetRootElement(m_hWnd, &rt);
		HTMLayoutUpdateElementEx(rt, MEASURE_INPLACE);
	}

	if( handled ) return lr; // HTMLayout processed the message
	return CWnd::WindowProc(message,wParam,lParam ); // call superclass window proc
}

IMPLEMENT_DYNAMIC(CHTMLayoutCtrl, CWnd)

BEGIN_MESSAGE_MAP(CHTMLayoutCtrl, CWnd)
END_MESSAGE_MAP()

CShapeSheetGrid* CHTMLayoutCtrl::FindGrid(LPCWSTR name) const
{
	for (int i = 0; i < m_impl->m_grids.GetSize(); ++i)
	{
		CShapeSheetGrid* grid = m_impl->m_grids[i];

		if (!lstrcmp(grid->GetName(), name))
			return grid;
	}

	return NULL;
}

const GridControls& CHTMLayoutCtrl::GetGridControls() const
{
	return m_impl->m_grids;
}
