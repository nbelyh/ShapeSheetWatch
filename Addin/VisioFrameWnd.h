
/*******************************************************************************

	@file 
	Declaration of "Visio frame window" - a "bridge" window between p4b and visio.

*******************************************************************************/

#pragma once

class CVisioFrameWnd 
	: public CWnd
{
public:
	CVisioFrameWnd();
	~CVisioFrameWnd();

	/// Constructs new Visio frame window.
	void Create(IVWindowPtr app);
	void Destroy();

protected:
	BOOL OnEraseBkgnd(CDC* pDC);
	void OnSize(UINT nType, int cx, int cy);
	BOOL OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);
	void OnDestroy();
	void PostNcDestroy();

	DECLARE_MESSAGE_MAP()
	void OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result);
	void OnDeleteItem(NMHDR*nmhdr, LRESULT* result);
	void OnContextMenu(CWnd* pWnd, CPoint point);

	void OnShowColumn(UINT cmd_id);
	void OnUpdateShowColumn(CCmdUI* pCmdUI);
	void OnInitMenuPopup(CMenu* pPopupMenu, UINT nIndex, BOOL bSysMenu);

private:
	struct Impl;
	Impl* m_impl;
};
