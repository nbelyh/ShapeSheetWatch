// Connect.cpp : Implementation of CConnect

#include "stdafx.h"

#include "Addin.h"
#include "AddIn_i.h"

#include "lib/PictureConvert.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
#include "lib/UI.h"
#include "lib/Language.h"

#include "Connect.h"

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

struct CConnect::Impl 
	: public VisioIdleTaskProcessor
	, public VEventHandler
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
		case (short)(Visio::visEvtApp|Visio::visEvtNonePending):
			OnVisioIdle();
			break;

		case (short)(Visio::visEvtApp|Visio::visEvtWinActivate):
			OnWindowActivated();
			break;

		case (short)(Visio::visEvtCodeWinOnAddonKeyMSG):
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
		Visio::IVApplicationPtr app;
		pApplication->QueryInterface(__uuidof(IDispatch), (LPVOID*)&app);

		pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_addin);

		if (GetVisioVersion(app) < 14)
			m_ui.CreateCommandBarsMenu(app);

		Visio::IVEventListPtr evt_list = 
			app->EventList;

		evt_idle.Advise(evt_list, Visio::visEvtApp|Visio::visEvtNonePending, this);
		evt_win_activated.Advise(evt_list, Visio::visEvtApp|Visio::visEvtWinActivate, this);
		evt_keystroke.Advise(evt_list, Visio::visEvtCodeWinOnAddonKeyMSG, this);

		theApp.SetVisioApp(app);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void Destroy() 
	{
		m_ui.DestroyCommandBarsMenu();

		evt_idle.Unadvise();
		evt_win_activated.Unadvise();
		evt_keystroke.Unadvise();

		theApp.SetVisioApp(NULL);

		ribbon  = NULL;
		m_addin = NULL;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	UINT GetControlCommand(IDispatch* pControl)
	{
		IRibbonControlPtr control;
		pControl->QueryInterface(__uuidof(IRibbonControl), (void**)&control);

		CComBSTR tag;
		if (FAILED(control->get_Tag(&tag)))
			return S_OK;

		return StrToInt(tag);
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

	VARIANT_BOOL IsRibbonButtonVisible(IDispatch * pControl)
	{
		return VARIANT_TRUE;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	VARIANT_BOOL IsRibbonButtonEnabled(IDispatch * pControl)
	{
		Visio::IVApplicationPtr app = theApp.GetVisioApp();

		Visio::IVDocumentPtr doc;
		if (FAILED(app->get_ActiveDocument(&doc)) || doc == NULL)
			return VARIANT_FALSE;

		Visio::VisDocumentTypes doc_type = Visio::visDocTypeInval;
		if (FAILED(doc->get_Type(&doc_type)) || doc_type == Visio::visDocTypeInval)
			return VARIANT_FALSE;

		return VARIANT_TRUE;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	VARIANT_BOOL IsRibbonButtonPressed(IDispatch * pControl)
	{
		Visio::IVApplicationPtr app = theApp.GetVisioApp();

		Visio::IVWindowPtr window;
		if (FAILED(app->get_ActiveWindow(&window)) || window == NULL)
			return VARIANT_FALSE;

		return theApp.GetWindowShapeSheet(GetVisioWindowHandle(window)) != NULL ? VARIANT_TRUE : VARIANT_FALSE;
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

	void SetRibbon(IDispatchPtr disp) 
	{
	}

	IDispatchPtr m_addin;

	Visio::IVApplicationPtr application;
	IDispatchPtr addin;
	IRibbonUIPtr ribbon;

	CVisioEvent	 evt_idle;
	CVisioEvent	 evt_win_activated;
	CVisioEvent	 evt_keystroke;

	CSimpleArray<VisioIdleTask*> idle_tasks;

	void AddVisioIdleTask(VisioIdleTask* task)
	{
		idle_tasks.Add(task);
	}

	void OnVisioIdle()
	{
		long idx = 0;
		while (idx < idle_tasks.GetSize())
		{
			VisioIdleTask* task = idle_tasks[idx];

			if (task->Execute())
			{
				idle_tasks.RemoveAt(idx);
				delete task;
			}
			else
			{
				++idx;
			} 
		}
	}

	void OnWindowActivated()
	{
		Office::IRibbonUIPtr ribbon = theApp.GetRibbon();
		
		if (ribbon)
			ribbon->Invalidate();
	}


	HRESULT OnKeystroke(Visio::IVMSGWrapPtr key_msg, VARIANT* pvResult)
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

		if (CWnd::FromHandlePermanent(msg.hwnd))
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

	void OnRibbonCheckboxClicked(IDispatch * pControl, VARIANT_BOOL * pvarfPressed)
	{
		UINT cmd_id = GetControlCommand(pControl);

		if (cmd_id > 0)
			theApp.OnCommand(cmd_id);
	}
};

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** custom)
{
	ENTER_METHOD()

	m_impl->Create(pApplication, pAddInInst);

	return S_OK;

	LEAVE_METHOD()
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD()

	m_impl->Destroy();
	return S_OK;

	LEAVE_METHOD()
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD();

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD();

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD();

	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
{
	ENTER_METHOD();

	return GetRibbonText(RibbonXml);

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnRibbonCheckboxClicked(IDispatch *pControl, VARIANT_BOOL *pvarfPressed)
{
	ENTER_METHOD()

	m_impl->OnRibbonCheckboxClicked(pControl, pvarfPressed);
	return S_OK;

	LEAVE_METHOD()
}

STDMETHODIMP CConnect::OnRibbonButtonClicked(IDispatch * disp)
{ 
	ENTER_METHOD();

	m_impl->OnRibbonButtonClicked(disp);
	return S_OK; 

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnRibbonLoad(IDispatch* disp)
{
	ENTER_METHOD();

	theApp.SetRibbon(disp);
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::OnRibbonLoadImage(BSTR bstrID, IPictureDisp ** ppdispImage)
{
	ENTER_METHOD();

	return CustomUiGetPng(MAKEINTRESOURCE(StrToInt(bstrID)), ppdispImage, NULL);

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::IsRibbonButtonPressed(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	*pResult = m_impl->IsRibbonButtonPressed(RibbonControl);
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::IsRibbonButtonEnabled(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	*pResult = m_impl->IsRibbonButtonEnabled(RibbonControl);
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::GetRibbonLabel(IDispatch *pControl, BSTR *pbstrLabel)
{
	ENTER_METHOD();

	*pbstrLabel = m_impl->GetRibbonLabel(pControl).AllocSysString();
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::GetRibbonImage(IDispatch *pControl, IPictureDisp ** ppdispImage)
{
	return S_OK;
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CConnect::IsRibbonButtonVisible(IDispatch * RibbonControl, VARIANT_BOOL* pResult)
{
	ENTER_METHOD();

	*pResult = m_impl->IsRibbonButtonVisible(RibbonControl);
	return S_OK;

	LEAVE_METHOD();
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

HRESULT CConnect::FinalConstruct ()
{
	return S_OK;
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

void CConnect::FinalRelease ()
{

}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

CConnect::CConnect ()
	: m_impl(new Impl())
{

}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

CConnect::~CConnect ()
{
	delete m_impl;
}
