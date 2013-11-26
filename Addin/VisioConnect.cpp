// Connect.cpp : Implementation of CVisioConnect

#include "stdafx.h"

#include "Addin.h"
#include "AddIn_i.h"

#include "lib/PictureConvert.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
#include "lib/UI.h"
#include "lib/Language.h"

#include "VisioConnect.h"

namespace {

	UINT GetControlCommand(IDispatch* pControl)
	{
		IRibbonControlPtr control;
		pControl->QueryInterface(__uuidof(IRibbonControl), (void**)&control);

		CComBSTR tag;
		if (FAILED(control->get_Tag(&tag)))
			return S_OK;

		return StrToInt(tag);
	}

}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

struct CVisioConnect::Impl 
	: public VEventHandler
	, IAddinUI
{
	AddinUi m_ui;

	HRESULT HandleVisioEvent(
		IN	IUnknown*	ipSink,			//	ipSink [assert]
		IN	short		nEventCode,		//	code of event that's firing.
		IN	IDispatch*	pSourceObj,		//	object that is firing event.
		IN	long		lEventID,		//	id of event that is firing.
		IN	long		lEventSeqNum,	//	sequence number of event.
		IN	IDispatch*	pSubjectObj,	//	subject of this event.
		IN	VARIANT		vMoreInfo,		//	other info.
		OUT VARIANT*	pvResult)		//	return a value to Visio for query events.
	{
		ENTER_METHOD();

		switch(nEventCode) 
		{
		case (short)(visEvtApp|visEvtIdle):
			theApp.ProcessIdleTasks();
			break;

		case (short)(visEvtApp|visEvtWinActivate):
			OnWindowActivated(pSubjectObj);
			break;

		case (short)(visEvtWindow|visEvtDel):
			OnWindowClosed(pSubjectObj);
			break;

		case (short)(visEvtWindow|visEvtAdd):
			OnWindowOpened(pSubjectObj);
			break;

		case (short)(visEvtCodeWinOnAddonKeyMSG):
			return OnKeystroke(pSubjectObj, pvResult);
			break;

		}

		return S_OK;

		LEAVE_METHOD();
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void Create(IDispatch * pApplication, IDispatch * pAddInInst) 
	{
		IVApplicationPtr app;
		pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&app);

		theApp.SetVisioApp(app);
		theApp.SetUI(this);

		pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_addin);

		if (GetVisioVersion() < 14)
			m_ui.CreateCommandBarsUI(app);

		IVEventListPtr evt_list = 
			app->EventList;

		evt_idle.Advise(evt_list, visEvtApp|visEvtIdle, this);
		evt_win_activated.Advise(evt_list, visEvtApp|visEvtWinActivate, this);
		evt_win_opened.Advise(evt_list, visEvtWindow|visEvtAdd, this);
		evt_win_closed.Advise(evt_list, visEvtWindow|visEvtDel, this);
		evt_keystroke.Advise(evt_list, visEvtCodeWinOnAddonKeyMSG, this);

		theApp.UpdateVisioUI();
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void Destroy() 
	{
		m_ui.DestroyCommandBarsUI();

		evt_idle.Unadvise();
		evt_win_activated.Unadvise();
		evt_win_opened.Unadvise();
		evt_win_closed.Unadvise();
		evt_keystroke.Unadvise();

		theApp.SetVisioApp(NULL);
		m_addin = NULL;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void OnRibbonButtonClicked(IDispatch * pControl) 
	{
		UINT cmd_id = GetControlCommand(pControl);

		if (cmd_id > 0)
			theApp.OnCommand(cmd_id);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	CString GetRibbonLabel(IDispatch* pControl) 
	{
		UINT cmd_id = GetControlCommand(pControl);

		LanguageLock lock(GetAppLanguage(theApp.GetVisioApp()));

		CString result;
		result.LoadString(cmd_id);

		return result;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	IDispatchPtr m_addin;

	CVisioEvent	 evt_idle;
	CVisioEvent	 evt_win_activated;
	CVisioEvent  evt_win_opened;
	CVisioEvent	 evt_win_closed;
	CVisioEvent	 evt_keystroke;

	void OnWindowOpened(IVWindowPtr window)
	{
		if (window != NULL && window->Type == visDrawing)
			theApp.ShowShapeSheetWatchWindow(window, theApp.GetViewSettings()->IsVisibleByDefault());

		theApp.UpdateVisioUI();
	}

	void OnWindowClosed(IVWindowPtr window)
	{
		if (window != NULL && window->Type == visDrawing)
			theApp.GetViewSettings()->SetVisibleByDefault(theApp.IsShapeSheetWatchWindowShown(window));
		theApp.UpdateVisioUI();
	}

	void OnWindowActivated(IVWindowPtr window)
	{
		theApp.UpdateVisioUI();
	}

	bool IsProcessedByAddon(WPARAM wp)
	{
		switch (wp)
		{
		case VK_DOWN:
		case VK_UP:
		case VK_LEFT:
		case VK_RIGHT:
		case VK_DELETE:
		case VK_TAB:
		case VK_NEXT:
		case VK_PRIOR:
		case VK_HOME:
		case VK_END:
		case VK_F2:
		case VK_ESCAPE:
		case VK_LBUTTON:
		case VK_RBUTTON:
		case VK_SPACE:
		case VK_RETURN:
			return true;
		}

		if (IsCTRLpressed())
		{
			CharLower((LPWSTR)wp);
			switch (wp)
			{
			case 'a':
			case 'c':
			case 'x':
			case 'k':
				return true;
			}
		}

		return false;
	}

	HRESULT OnKeystroke(IVMSGWrapPtr key_msg, VARIANT* pvResult)
	{
		ASSERT_RETURN_VALUE(key_msg != NULL, E_INVALIDARG);
		ASSERT_RETURN_VALUE(pvResult != NULL, E_INVALIDARG);

		MSG msg;

		msg.hwnd    = reinterpret_cast<HWND>(key_msg->Gethwnd());
		msg.message = key_msg->Getmessage();
		msg.wParam  = key_msg->GetwParam();
		msg.lParam  = key_msg->GetlParam();
		msg.pt.x    = key_msg->Getptx();
		msg.pt.y    = key_msg->Getpty();
		msg.time    = key_msg->Getposttime();

		bool result = false;

		if (CWnd::FromHandlePermanent(msg.hwnd) && IsProcessedByAddon(msg.wParam))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
			result = true;
		}
		else
		{
			result = CWnd::WalkPreTranslateTree(NULL, &msg) != FALSE;
		}

		*pvResult = variant_t(result);
		return S_OK;
	}

	virtual void Update()
	{
		if (m_ribbon)
			m_ribbon->Invalidate();
		else
			m_ui.UpdateCommandBarsUI();
	}

	Office::IRibbonUIPtr m_ribbon;
};

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** custom)
{
	ENTER_METHOD()

	m_impl->Create(pApplication, pAddInInst);

	theApp.GetViewSettings()->Load();
	return S_OK;

	LEAVE_METHOD()
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD()

	m_impl->Destroy();

	theApp.GetViewSettings()->Save();
	return S_OK;

	LEAVE_METHOD()
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD();

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD();

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD();

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
{
	ENTER_METHOD();

	return GetRibbonText(RibbonXml);

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnRibbonButtonClicked(IDispatch * disp)
{ 
	ENTER_METHOD();

	m_impl->OnRibbonButtonClicked(disp);
	return S_OK; 

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnRibbonLoad(IDispatch* disp)
{
	ENTER_METHOD();

	return disp->QueryInterface(__uuidof(Office::IRibbonUI), (void**)&m_impl->m_ribbon);

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnRibbonLoadImage(BSTR bstrID, IPictureDisp ** ppdispImage)
{
	ENTER_METHOD();

	return CustomUiGetPng(MAKEINTRESOURCE(StrToInt(bstrID)), ppdispImage, NULL);

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::IsRibbonButtonEnabled(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	*pResult = theApp.IsCommandEnabled(GetControlCommand(RibbonControl));
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::GetRibbonLabel(IDispatch *pControl, BSTR *pbstrLabel)
{
	ENTER_METHOD();

	*pbstrLabel = m_impl->GetRibbonLabel(pControl).AllocSysString();
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::GetRibbonImage(IDispatch *pControl, IPictureDisp ** ppdispImage)
{
	return S_OK;
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::IsRibbonButtonVisible(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	*pResult = theApp.IsCommandVisible(GetControlCommand(RibbonControl));
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnRibbonCheckboxClicked(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	theApp.OnCommand(GetControlCommand(RibbonControl));

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::IsRibbonButtonPressed(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	*pResult = theApp.IsCommandChecked(GetControlCommand(RibbonControl));
	return S_OK;

	LEAVE_METHOD();
}


/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

HRESULT CVisioConnect::FinalConstruct ()
{
	return S_OK;
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

void CVisioConnect::FinalRelease ()
{

}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

CVisioConnect::CVisioConnect ()
	: m_impl(new Impl())
{

}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

CVisioConnect::~CVisioConnect ()
{
	delete m_impl;
}
