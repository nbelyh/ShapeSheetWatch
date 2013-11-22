//	Copyright (C) Microsoft Corporation. All rights reserved.
//	<summary>
//	handling.
//	</summary>

#include "stdafx.h"
#include "Visio.h"

extern "C" {

/**-----------------------------------------------------------------------------
	Match the new IVisEventProc signature verbatim except that the
	first param here is the sink object that is calling the callback
	proc. (An additional param over what's received in VisEventProc.)
------------------------------------------------------------------------------*/

typedef HRESULT (STDMETHODCALLTYPE VISEVENTPROC)(
		/* [in] */ IUnknown* ipSink,			//	ipSink [assert]
		/* [in] */ short nEventCode,			//	code of event that's firing.
		/* [in] */ IDispatch* pSourceObj,		//	object that is firing event.
		/* [in] */ long nEventID,				//	id of event that is firing.
		/* [in] */ long nEventSeqNum,			//	sequence number of event.
		/* [in] */ IDispatch* pSubjectObj,		//	subject of this event.
		/* [in] */ VARIANT vMoreInfo,			//	other info.
		/* [retval][out] */ VARIANT* pvResult);	//	return a value to Visio for query events.

typedef VISEVENTPROC *LPVISEVENTPROC;

HRESULT CoCreateAddonSink(LPVISEVENTPROC pCallback, IUnknown FAR* FAR* ppSink);

}	//	end of extern "C"

//	To allow compiling with DevStudio pre-6.0:
#ifndef LOCALE_NEUTRAL
#define LOCALE_NEUTRAL	MAKELCID(MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL), SORT_DEFAULT)
#endif

/**-----------------------------------------------------------------------------
CVisioAddonSink - Sink class that Visio uses to notify about events

This class implements the IVisEventProc event notification interface.
Use this class when signing up for Visio events using the AddAdvise
method.
------------------------------------------------------------------------------*/

class CVisioAddonSink : public IVisEventProc
{
public:
	//	IUnknown methods

	STDMETHOD(QueryInterface)(REFIID riid, void FAR* FAR* ppv);
	STDMETHOD_(ULONG, AddRef)(void);
	STDMETHOD_(ULONG, Release)(void);

	//	IDispatch methods

	STDMETHOD(GetTypeInfoCount)(
		UINT FAR* pctinfo);

	STDMETHOD(GetTypeInfo)( 
		UINT itinfo,
		LCID lcid,
		ITypeInfo FAR* FAR* pptinfo);

	STDMETHOD(GetIDsOfNames)( 
		REFIID riid,
		LPOLESTR FAR* rgszNames,
		UINT cNames,
		LCID lcid,
		DISPID FAR* rgdispid);

	STDMETHOD(Invoke)( 
		DISPID dispidMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS FAR* pdispparams,
		VARIANT FAR* pvarResult,
		EXCEPINFO FAR* pexcepinfo,
		UINT FAR* puArgErr);

	//	IVisEventProc method

	virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE raw_VisEventProc( 
		/* [in] */ short nEventCode,
		/* [in] */ IDispatch __RPC_FAR* pSourceObj,
		/* [in] */ long nEventID,
		/* [in] */ long nEventSeqNum,
		/* [in] */ IDispatch __RPC_FAR* pSubjectObj,
		/* [in] */ VARIANT vMoreInfo,
		/* [retval][out] */ VARIANT __RPC_FAR* pvResult);


protected:

	/**-----------------------------------------------------------------------------
	CVisioAddonSink objects and derived objects must be created with
	a ref count and destroyed by a final Release. This is the reason
	behind having the CoCreate-style helper functions. Visio *WILL*
	unload VSL's first at shutdown time and then Release all outstanding
	event sinks. That's why it's important to call CVisioEvent::Delete
	at VSL unload time... That forces Visio to Release your sink
	objects while they're still viable.
	------------------------------------------------------------------------------*/

	//	Constructor - should only be used for sub-classing purposes.
	//	Subclass should override VisEventProc. This constructor leaves
	//	both m_pCallback and m_pHandler NULL.
	CVisioAddonSink();

	//	Protected destructor enforces Release as delete mechanism.
	virtual ~CVisioAddonSink();

private:
	//	Unimplemented private copy constructor.
	CVisioAddonSink(const CVisioAddonSink&);

	//	Other constructors ONLY accessible through friend functions.
	friend HRESULT CoCreateAddonSink(
		LPVISEVENTPROC pCallback,
		IUnknown FAR* FAR* ppSink);

	friend HRESULT CoCreateAddonSinkForHandler(
		LPVISEVENTPROC pCallback, 
		VEventHandler* pHandler, 
		IUnknown FAR*  FAR* ppSink);

	CVisioAddonSink(LPVISEVENTPROC pCallback, VEventHandler* pHandler= NULL);

	//	Common initialization called only from constructors:
	void CommonConstruct(
		LPVISEVENTPROC pCallback= NULL,
		VEventHandler* pHandler= NULL);

	//	Data members:
	static ITypeInfo FAR* m_pInfo;	//	IDispatch methods fail gracefully if NULL
	ULONG m_cRef;					//	reference count
	LPVISEVENTPROC m_pCallback;		//	function to call when VisEventProc gets called...
	VEventHandler *m_pHandler;		//	object to call when VisEventProc gets called...
};

//	Type info pointer shared across all instances of CVisioAddonSink:
ITypeInfo FAR* CVisioAddonSink::m_pInfo = NULL;

//	CommonConstruct - Common contructor code
//	<summary>
//	This method contains common constructor code for create CVisioAddonSink 
//	objects.
//	</summary>
//	Return Value:
//	None
void CVisioAddonSink::CommonConstruct(LPVISEVENTPROC pCallback /*=NULL*/,
									  VEventHandler* pHandler /*=NULL*/)
{
	HRESULT hr = NOERROR;

	if (m_pInfo)
	{
		m_pInfo->AddRef();
	}
	else
	{
		//	Visio 2000 was the first version of Visio to declare IVisEventProc
		//	in its type library. The major.minor version number of the Visio
		//	2000 type library is 4.7.

		ITypeLib* ipLib = NULL;

		hr = LoadRegTypeLib(__uuidof(__Visio), 4, 7, LOCALE_NEUTRAL, &ipLib);
		if (SUCCEEDED(hr))
		{
			hr = ipLib->GetTypeInfoOfGuid(__uuidof(IVisEventProc), &m_pInfo);

#ifdef _DEBUG
			if (FAILED(hr))
			{
				OutputDebugString(L"WARNING: Couldn't load type info for IVisEventProc.");
			}
#endif

			ipLib->Release();
			ipLib = NULL;
		}
	}

	m_cRef = 0;
	m_pCallback = pCallback;
	m_pHandler = pHandler;
}

//	CVisioAddonSink - Default constructor
//	<summary>
//	This method is the default constructor the the CVisioAddonSink class.
//	</summary>
//	Return Value:
//	None

CVisioAddonSink::CVisioAddonSink()
{
	this->CommonConstruct();
}

//	CVisioAddonSink - Constructor
//	<summary>
//	This method constructs CVisioAddonSink objects given a call back method and
//	a pointer to a VEventHandler interface.
//	</summary>
//	Return Value:
//	None

CVisioAddonSink::CVisioAddonSink(LPVISEVENTPROC pCallback,
								 VEventHandler* pHandler)
{
	this->CommonConstruct(pCallback, pHandler);
}

//	~CVisioAddonSink - Destructor
//	<summary>
//	This is the destructor method for the CVisioAddonSink class.
//	</summary>
//	Return Value:
//	None

CVisioAddonSink::~CVisioAddonSink()
{
	if ( m_pInfo )
	{
		if (0 == m_pInfo->Release())
		{
			//	If this is the last instance using m_pInfo then set 
			//	it to NULL so that we don’t leave a dangling static pointer. 

			m_pInfo = NULL;
		}
	}
}

//	<summary>
//	This method implements the IUnknown QueryInterface method for the 
//	CVisioAddonSink class.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code

STDMETHODIMP CVisioAddonSink::QueryInterface(REFIID riid,
											 void FAR* FAR* ppv)
{
	if (IID_IUnknown == riid ||
		IID_IDispatch == riid ||
		__uuidof(IVisEventProc) == riid)
	{
		*ppv = this;
		this->AddRef();
		return NOERROR;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

//	<summary>
//	This method implements the IUnknown AddRef method for the CVisioAddonSink 
//	class.
//	</summary>
//	Return Value:
//	ULONG - reference count after AddRef has finished

STDMETHODIMP_(ULONG) CVisioAddonSink::AddRef(void)
{
	m_cRef++;

	return m_cRef;
}

//	<summary>
//	This method implements the IUnknown Release method for the CVisioAddonSink 
//	class.
//	</summary>
//	Return Value:
//	ULONG - reference count after Release has finished

STDMETHODIMP_(ULONG) CVisioAddonSink::Release(void)
{
	m_cRef--;

	if (m_cRef == 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return m_cRef;
	}
}

//	<summary>
//	This method implements the IDispatch GetTypeInfoCount method for the 
//	CVisioAddonSink class.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR

STDMETHODIMP CVisioAddonSink::GetTypeInfoCount(UINT FAR* pctinfo)
{
	*pctinfo = 1;
	return NOERROR;
}

//	<summary>
//	This method implements the IDispatch GetTypeInfo method for the 
//	CVisioAddonSink class.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code

STDMETHODIMP CVisioAddonSink::GetTypeInfo(
    /* [in] */ UINT itinfo,
    /* [in] */ LCID /*lcid*/,
    /* [out] */ ITypeInfo FAR* FAR* pptinfo)
{
	HRESULT hr = NOERROR;

	if ( pptinfo && m_pInfo )
	{
		if ( itinfo == 0 )
		{
			m_pInfo->AddRef();
			*pptinfo = m_pInfo;
		}
		else
		{
			*pptinfo = NULL;
			hr = DISP_E_BADINDEX;
		}
	}
	else
	{
		hr = DISP_E_EXCEPTION;
	}

	return hr;
}

//	<summary>
//	This method implements the IDispatch GetIDsOfNames method for the 
//	CVisioAddonSink class.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code

STDMETHODIMP CVisioAddonSink::GetIDsOfNames(
    /* [in] */ REFIID /*riid*/,
    /* [size_is][in] */ LPOLESTR FAR* rgszNames,
    /* [in] */ UINT cNames,
    /* [in] */ LCID /*lcid*/,
    /* [size_is][out][in] */ DISPID FAR* rgdispid)
{
	HRESULT hr;

	if ( m_pInfo )
	{
		hr = DispGetIDsOfNames(
			m_pInfo, 
			rgszNames, 
			cNames, 
			rgdispid);
	}
	else
	{
		hr = DISP_E_EXCEPTION;
	}

	return hr;
}

//	<summary>
//	This method implements the IDispatch Invoke method for the 
//	CVisioAddonSink class.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code

STDMETHODIMP CVisioAddonSink::Invoke(
    /* [in] */ DISPID dispidMember,
    /* [in] */ REFIID /*riid*/,
    /* [in] */ LCID /*lcid*/,
    /* [in] */ WORD wFlags,
    /* [unique][in] */ DISPPARAMS FAR* pdispparams,
    /* [unique][out][in] */ VARIANT FAR* pvarResult,
    /* [out] */ EXCEPINFO FAR* pexcepinfo,
    /* [out] */ UINT FAR* puArgErr)
{
	HRESULT hr;

	if ( m_pInfo )
	{
		hr = DispInvoke(this,
					   m_pInfo,
					   dispidMember,
					   wFlags,
					   pdispparams,
					   pvarResult,
					   pexcepinfo,
					   puArgErr);
	}
	else
	{
		hr = DISP_E_EXCEPTION;
	}

	return hr;
}

//	<summary>
//	This method handles the event callback received from Visio when an event
//	takes place.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code

//	Using #import for wrapping Visio API
HRESULT STDMETHODCALLTYPE CVisioAddonSink::raw_VisEventProc(
    /* [in] */ short nEventCode,
    /* [in] */ IDispatch __RPC_FAR* pSourceObj,
    /* [in] */ long nEventID,
    /* [in] */ long nEventSeqNum,
    /* [in] */ IDispatch __RPC_FAR* pSubjectObj,
    /* [in] */ VARIANT vMoreInfo,
    /* [retval][out] */ VARIANT __RPC_FAR* pvResult)
{
	//	If default constructor was used (say CVisioAddonSink was subclassed)
	//	then m_pCallback will be 0.
	//	In this case VisEventProc must be overridden by the subclass.

	HRESULT hr = NOERROR;

	if (m_pHandler)
	{
		hr = m_pHandler->HandleVisioEvent(this,
			                             nEventCode,
										 pSourceObj,
										 nEventID,
										 nEventSeqNum,
										 pSubjectObj,
										 vMoreInfo,
										 pvResult);
	}

	if (m_pCallback && SUCCEEDED(hr))
	{
		hr = (*m_pCallback)(this,
						   nEventCode,
						   pSourceObj,
						   nEventID,
						   nEventSeqNum,
						   pSubjectObj,
						   vMoreInfo,
						   pvResult);
	}

	return hr;
}

//	CoCreateAddonSink - Create a CVisioAddonSink object
//	<summary>
//	This function creates a CVisioAddonSink object to use with the AddAdvise
//	method. The method takes a pointer to a function with the VISEVENTPROC
//	signature.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code
HRESULT CoCreateAddonSink(LPVISEVENTPROC pCallback,
						  IUnknown FAR* FAR* ppSink)
{
	HRESULT hr = E_FAIL;

	//	assert ppSink != NULL

	*ppSink = new CVisioAddonSink(pCallback);

	if (*ppSink)
	{
		(*ppSink)->AddRef();
		hr = NOERROR;
	}

	return hr;
}

//	CoCreateAddonSinkForHandler - Create a CVisioAddonSink object
//	<summary>
//	This function creates a CVisioAddonSink object to use with the AddAdvise
//	method.  The method takes a pointer to a function with the VISEVENTPROC
//	signature and a pointer to an object that implements the VEventHandler 
//	interface.
//	</summary>
//	Return Value:
//	HRESULT - NOERROR or some HRESULT error code
HRESULT CoCreateAddonSinkForHandler(LPVISEVENTPROC pCallback,
									VEventHandler* pHandler,
									IUnknown FAR* FAR* ppSink)
{
	HRESULT hr = E_FAIL;

	//	assert ppSink != NULL

	*ppSink = new CVisioAddonSink(pCallback, pHandler);

	if (*ppSink)
	{
		(*ppSink)->AddRef();
		hr = NOERROR;
	}

	return hr;
}

/**-----------------------------------------------------------------------------
	Event connector constructor
------------------------------------------------------------------------------*/

CVisioEvent::CVisioEvent ()
{
}

/**-----------------------------------------------------------------------------
	Automatically disconnects the event on destruction
------------------------------------------------------------------------------*/

CVisioEvent::~CVisioEvent()
{
	Unadvise();
}

/**-----------------------------------------------------------------------------
	Attaches given handler interface to given visio event.
	The event object generated is stored for detaching later.
------------------------------------------------------------------------------*/

HRESULT CVisioEvent::Advise(IVEventListPtr evt_list,
						 int evt_code,
						 VEventHandler* target)
{
	variant_t v_sink;
	
	V_VT(&v_sink)		= VT_UNKNOWN;
	V_UNKNOWN(&v_sink)	= 0;

	HRESULT hr = CoCreateAddonSinkForHandler(NULL, target, &V_UNKNOWN(&v_sink));

	if (FAILED(hr))
		return hr;

	m_event = evt_list->AddAdvise(static_cast<short>(evt_code), variant_t(v_sink), "", "");
	return S_OK;
}

/**-----------------------------------------------------------------------------
	Detaches this event from it's event source.
------------------------------------------------------------------------------*/

void CVisioEvent::Unadvise()
{
	if (m_event != NULL)
	{
		m_event->Delete();
		m_event.Release();
	}
}

/**-----------------------------------------------------------------------------
	Given visio window, returns HWND (window handle) for it.
------------------------------------------------------------------------------*/

HWND GetVisioWindowHandle (IVWindowPtr visio_window)
{
	return reinterpret_cast<HWND>(visio_window->GetWindowHandle32());
}

HWND GetVisioAppWindowHandle (IVApplicationPtr app)
{
	return reinterpret_cast<HWND>(app->GetWindowHandle32());
}

/**-----------------------------------------------------------------------
	Helper structure to lock several modeler action 
	in single undo scope (undo/redo support class)
------------------------------------------------------------------------*/

int VisioScopeLock::g_level = 0;

VisioScopeLock::VisioScopeLock(IVApplicationPtr app, LPCWSTR description)
	: m_result(VARIANT_FALSE)
	, m_scope_id(0)
	, m_app(app)
{
	if (g_level == 0)
		m_scope_id = m_app->BeginUndoScope(description);

	++g_level;
}

VisioScopeLock::~VisioScopeLock()
{
	if (m_scope_id != 0)
		m_app->EndUndoScope(m_scope_id, m_result);

	--g_level;
}

void VisioScopeLock::Commit()
{
	m_result = VARIANT_TRUE;
}

bool VisioScopeLock::IsInVisioScopeLock ()
{
	return g_level > 0;
}
