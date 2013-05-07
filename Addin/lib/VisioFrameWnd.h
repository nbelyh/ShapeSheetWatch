
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
	void Create(Visio::IVWindowPtr app);
	void Destroy();

protected:
	BOOL OnEraseBkgnd(CDC* pDC);
	void OnSize(UINT nType, int cx, int cy);
	BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	void OnDestroy();

	DECLARE_MESSAGE_MAP()
private:
	struct Impl;
	Impl* m_impl;
};
