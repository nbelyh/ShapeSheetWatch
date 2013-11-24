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

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

struct CVisioConnect::Impl 
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
		case (short)(visEvtApp|visEvtNonePending):
			OnVisioIdle();
			break;

		case (short)(visEvtApp|visEvtWinActivate):
			OnWindowActivated();
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

		pAddInInst->QueryInterface(__uuidof(IDispatch), (LPVOID*)&m_addin);

		if (GetVisioVersion() < 14)
			m_ui.CreateCommandBarsUI(app);

		IVEventListPtr evt_list = 
			app->EventList;

		evt_idle.Advise(evt_list, visEvtApp|visEvtNonePending, this);
		evt_win_activated.Advise(evt_list, visEvtApp|visEvtWinActivate, this);
		evt_keystroke.Advise(evt_list, visEvtCodeWinOnAddonKeyMSG, this);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void Destroy() 
	{
		m_ui.DestroyCommandBarsUI();

		evt_idle.Unadvise();
		evt_win_activated.Unadvise();
		evt_keystroke.Unadvise();

		theApp.SetVisioApp(NULL);
		m_ribbon = NULL;
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
		IVApplicationPtr app = theApp.GetVisioApp();

		IVDocumentPtr doc;
		if (FAILED(app->get_ActiveDocument(&doc)) || doc == NULL)
			return VARIANT_FALSE;

		VisDocumentTypes doc_type = visDocTypeInval;
		if (FAILED(doc->get_Type(&doc_type)) || doc_type == visDocTypeInval)
			return VARIANT_FALSE;

		return VARIANT_TRUE;
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
	IRibbonUIPtr m_ribbon;

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
		if (m_ribbon)
			m_ribbon->Invalidate();
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

	void SetRibbon(IDispatch* disp)
	{
		m_ribbon = disp;
	}
};

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** custom)
{
	ENTER_METHOD()

	m_impl->Create(pApplication, pAddInInst);

	return S_OK;

	LEAVE_METHOD()
}

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	ENTER_METHOD()

	m_impl->Destroy();
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

	m_impl->SetRibbon(disp);
	return S_OK;

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

	*pResult = m_impl->IsRibbonButtonEnabled(RibbonControl);
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

	*pResult = m_impl->IsRibbonButtonVisible(RibbonControl);
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
