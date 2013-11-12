
/*******************************************************************************

	@file 

*******************************************************************************/

#pragma once

#include "lib/HTMLayoutCtrl/HTMLayoutCtrl.h"

class CVisioFrameWnd 
	: public CHTMLayoutCtrl
{
public:
	void Create(IVWindowPtr app);
	void Destroy();

protected:
	BOOL OnEraseBkgnd(CDC* pDC);
	void OnSize(UINT nType, int cx, int cy);
	BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	void PostNcDestroy();

	DECLARE_MESSAGE_MAP()

	CHTMLayoutCtrl m_html;
};
