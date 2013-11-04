
/*******************************************************************************
	ADDSINK.H - Definitions for client side of Visio Advise Sink.
	Copyright (C) Microsoft Corporation. All rights reserved.

	This header file contains the prototype for the event call back
	function (VISEVENTPROC) along with the class definitions for VEventHandler 
	and CVisioAddonSink.

*******************************************************************************/

#pragma once

HWND GetVisioWindowHandle (IVWindowPtr visio_window);
HWND GetVisioAppWindowHandle (IVApplicationPtr app);

/**-----------------------------------------------------------------------------
	VEventHandler - definition of interface for handling Visio events

	This abstract base class contains a single method through which event
	notification will take place. Inherit your C++ class from this Class if 
	you want a C++ object context (a this pointer) when you get called by 
	VisEventProc in response to a Visio event.
------------------------------------------------------------------------------*/

class VEventHandler
{

public:

	/**-----------------------------------------------------------------------------
		This signature for HandleVisioEvent matches the new IVisEventProc 
		signature verbatim except that the first param here is the 
		sink object that is calling the callback proc. (An additional param
		over what's received in VisEventProc.)
	------------------------------------------------------------------------------*/

	virtual HRESULT HandleVisioEvent(
		/* [in] */ IUnknown* ipSink,			//	ipSink [assert]
        /* [in] */ short nEventCode,			//	code of event that's firing.
        /* [in] */ IDispatch* pSourceObj,		//	object that is firing event.
        /* [in] */ long nEventID,				//	id of event that is firing.
        /* [in] */ long nEventSeqNum,			//	sequence number of event.
        /* [in] */ IDispatch* pSubjectObj,		//	subject of this event.
        /* [in] */ VARIANT vMoreInfo,			//	other info.
        /* [retval][out] */ VARIANT* pvResult	//	return a value to Visio 
												//	for query events.
	) = 0;

};

/**-----------------------------------------------------------------------------
	Helper class to help advising to visio events.
	Provides the functionality to connect or disconnect a class 
	that implements VEventHander interface to given event.
------------------------------------------------------------------------------*/

struct CVisioEvent
{
public:
	CVisioEvent	();
	~CVisioEvent();

	HRESULT Advise(IVEventListPtr evt_list, int evt_code, VEventHandler* target);
	void Unadvise();
private:
	IVEventPtr m_event;
};

/**---------------------------------------------------------------------------------
	
-----------------------------------------------------------------------------------*/

struct VisioScopeLock
{
	VisioScopeLock(IVApplicationPtr app, LPCWSTR description);
	~VisioScopeLock();

	static bool IsInVisioScopeLock();
	void Commit();

private:
	IVApplicationPtr m_app;
	long			m_scope_id;
	VARIANT_BOOL	m_result;

	static int		g_level;
};

/**---------------------------------------------------------------------------------
	
-----------------------------------------------------------------------------------*/

struct VisioIdleTask
{
	virtual bool Execute() = 0;
};

struct VisioIdleTaskProcessor
{
	virtual void AddVisioIdleTask(VisioIdleTask* task) = 0;
};


/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

int GetVisioVersion(IVApplicationPtr app);
