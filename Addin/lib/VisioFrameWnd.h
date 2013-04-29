
/*******************************************************************************

	@file 
	Declaration of "Visio frame window" - a "bridge" window between p4b and visio.

*******************************************************************************/

#pragma once
#include "lib/HTMLayoutCtrl.h"

class CVisioFrameWnd 
	: public CWnd
{
public:
	CVisioFrameWnd();
	~CVisioFrameWnd();

	/// Constructs new Visio frame window.
	virtual void Create(Visio::IVApplicationPtr app);

protected:
	BOOL OnEraseBkgnd(CDC* pDC);
	void OnSize(UINT nType, int cx, int cy);
	void OnDestroy();

	DECLARE_MESSAGE_MAP()

private:
	struct Impl;
	Impl* m_impl;
};
