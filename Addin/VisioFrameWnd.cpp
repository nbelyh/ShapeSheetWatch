
/*******************************************************************************

	@file Implementation of a Visio frame window.

*******************************************************************************/

#include "stdafx.h"
#include "VisioFrameWnd.h"
#include "lib/GridCtrl/GridCtrl.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
#include "ShapeSheet.h"
#include "Addin.h"

/**-----------------------------------------------------------------------------
	Message map
------------------------------------------------------------------------------*/

BEGIN_MESSAGE_MAP(CVisioFrameWnd, CWnd)
	ON_WM_CREATE()
	ON_WM_DESTROY()
	ON_WM_ERASEBKGND()
	ON_WM_SIZE()
	ON_WM_MOUSEWHEEL()
	ON_NOTIFY(GVN_ENDLABELEDIT, 1, OnEndLabelEdit)
	ON_NOTIFY(GVN_DELETEITEM, 1, OnDeleteItem)
	ON_WM_CONTEXTMENU()
	ON_WM_INITMENUPOPUP()

	ON_COMMAND_RANGE(ID_ShowColumn, ID_ShowColumn+Column_Count, OnShowColumn)
	ON_UPDATE_COMMAND_UI_RANGE(ID_ShowColumn, ID_ShowColumn+Column_Count, OnUpdateShowColumn)
END_MESSAGE_MAP()

#define COLOR_TH_BK	RGB(153,180,209)
#define COLOR_TH_FG	RGB(0,0,0)

#define COLOR_SRC_BK	RGB(0xE3,0xE3,0xE3)
#define COLOR_SRC_FG	RGB(0xF0,0,0)

struct CVisioFrameWnd::Impl : public VEventHandler
{
	virtual HRESULT HandleVisioEvent(
		IN      IUnknown*       ipSink,
		IN      short           nEventCode,
		IN      IDispatch*      pSourceObj,
		IN      long            lEventID,
		IN      long            lEventSeqNum,
		IN      IDispatch*      pSubjectObj,
		IN      VARIANT         vMoreInfo,
		OUT VARIANT*    pvResult)
	{
		ENTER_METHOD();

		switch(nEventCode) 
		{
		case (short)(visEvtCodeWinSelChange):
			OnSelectionChanged(pSubjectObj);
			break;

		case (short)(visEvtMod|visEvtCell):
			OnCellChanged(pSubjectObj);
			break;

		case (short)(visEvtCodeBefSelDel):
			OnSelectionDelete(pSubjectObj);
			break;
		}

		return S_OK;

		LEAVE_METHOD()
	}

	CString GetColumnName(int i)
	{
		switch (i)
		{
		case Column_Mask:		return L"Mask";
		case Column_Name:		return L"Name";
		case Column_S:			return L"Section";
		case Column_R:			return L"Row";
		case Column_RU:			return L"Row (U)";
		case Column_C:			return L"Column";
		case Column_Formula:	return L"Formula";
		case Column_FormulaU:	return L"Formula (U)";
		case Column_Value:		return L"Result";
		case Column_ValueU:		return L"Result (U)";

		default:	return L"";
		}
	}

	void UpdateGridColumns()
	{
		grid.SetFixedRowCount(1);

		grid.SetColumnCount(Column_Count);

		for (int i = 0; i < Column_Count; ++i)
		{
			grid.SetItemBkColour(0, i, COLOR_TH_BK);
			grid.SetItemFgColour(0, i, COLOR_TH_FG);

			grid.SetItemText(0, i, GetColumnName(i));

			if (m_view_settings->IsColumnHidden(i))
				grid.SetColumnWidth(i, 0);
		}
	}

	CVisioEvent m_evt_shape_deleted;
	CVisioEvent m_evt_cell_changed;
	CVisioEvent	m_evt_sel_changed;

	void Create(IVWindowPtr window)
	{
		IVEventListPtr event_list = window->GetEventList();
		m_evt_sel_changed.Advise(event_list, visEvtCodeWinSelChange, this);

		IVEventListPtr doc_event_list = window->GetDocument()->GetEventList();
		m_evt_shape_deleted.Advise(doc_event_list, visEvtCodeBefSelDel, this);

		m_view_settings = theApp.GetViewSettings();

		UpdateGridColumns();
		UpdateGridRows();
	}

	void Destroy()
	{
		m_evt_sel_changed.Unadvise();
		m_evt_shape_deleted.Unadvise();

		grid.DeleteAllItems();
	}

	void OnSelectionDelete(IVSelectionPtr selection)
	{
		if (m_shape && selection)
		{
			for (long i = 1; i <= selection->Count; ++i)
			{
				IVShapePtr shape = selection->Item[i];

				if (shape->ID == m_shape->ID)
					SetShape(NULL);
			}
		}
	}

	void OnCellChanged(IVCellPtr cell)
	{
		UpdateGridRows();
	}

	IVShapePtr m_shape;

	void SetShape(IVShapePtr shape)
	{
		m_evt_cell_changed.Unadvise();

		if (shape)
		{
			IVEventListPtr event_list = shape->GetEventList();
			m_evt_cell_changed.Advise(event_list, (visEvtMod|visEvtCell), this);
		}

		m_shape = shape;
	}

	void OnSelectionChanged(IVWindowPtr window)
	{
		IVSelectionPtr selection = window->Selection;

		IVShapePtr shape = (selection->Count == 1)
			? selection->Item[1] : NULL;
				
		SetShape(shape);

		UpdateGridRows();
	}

	void SetHeadColumn(short row, short col, const CString& text)
	{
		grid.SetItemBkColour(row, col, COLOR_SRC_BK);
		grid.SetItemFgColour(row, col, COLOR_SRC_FG);
		grid.SetItemText(row, col, text);
	}

	ViewSettings* m_view_settings;

	void UpdateGridRows()
	{
		const Strings& cell_name_masks = m_view_settings->GetCellMasks();

		typedef std::vector< std::vector<SRC> > GroupCellInfos;
		GroupCellInfos cell_names(cell_name_masks.size());

		if (m_shape != NULL)
		{
			for (size_t i = 0; i < cell_name_masks.size(); ++i)
				GetCellNames(m_shape, cell_name_masks[i], cell_names[i]);
		}

		grid.SetRowCount(1);

		size_t row_count = 0;
		for (size_t i = 0; i < cell_names.size(); ++i)
			row_count += cell_names[i].size() > 0 ? cell_names[i].size() : 1;

		grid.SetRowCount(1 + row_count);

		size_t row = 1;
		for (size_t i = 0; i < cell_name_masks.size(); ++i)
		{
			SetHeadColumn(row, Column_Mask, cell_name_masks[i]);

			size_t m_row = row;

			if (cell_names[i].empty())
			{
				grid.SetItemData(row, Column_Mask, i);

				SetHeadColumn(row, Column_S, L"");
				SetHeadColumn(row, Column_R, L"");
				SetHeadColumn(row, Column_C, L"");

				++row;
			}
			else 
			{
				size_t s_start = row;
				size_t s_count = 0;
				short s_last = -1;

				size_t r_start = row;
				size_t r_count = 0;
				short r_last = -1;

				for (size_t j = 0; j < cell_names[i].size(); ++j)
				{
					SRC& src = cell_names[i][j];

					SetHeadColumn(row, Column_Name, src.name);

					SetHeadColumn(row, Column_S, src.s_name);
					SetHeadColumn(row, Column_R, src.r_name_l);
					SetHeadColumn(row, Column_RU, src.r_name_u);
					SetHeadColumn(row, Column_C, src.c_name);

					IVCellPtr cell = m_shape->GetCellsSRC(src.s, src.r, src.c);

					grid.SetItemData(row, Column_Mask, i);

					grid.SetItemText(row, Column_Formula, cell->Formula);
					grid.SetItemText(row, Column_FormulaU, cell->Formula);

					grid.SetItemText(row, Column_Value, cell->ResultStr[-1]);
					grid.SetItemText(row, Column_ValueU, cell->ResultStrU[-1]);

					if (s_last == src.s)
						++s_count;
					else
					{
						if (s_count > 1)
							grid.MergeCells(s_start, Column_S, row - 1, Column_S);

						s_last = src.s;
						s_start = row;
						s_count = 1;
					}

					if (r_last == src.r)
						++r_count;
					else
					{
						if (r_count > 1)
							grid.MergeCells(r_start, Column_R, row - 1, Column_R);

						r_last = src.r;
						r_start = row;
						r_count = 1;
					}

					++row;
				}

				if (s_count > 1)
					grid.MergeCells(s_start, Column_S, row - 1, Column_S);

				if (r_count > 1)
					grid.MergeCells(r_start, Column_R, row - 1, Column_R);
			}

			if (row > m_row + 1)
				grid.MergeCells(m_row, Column_Mask, row - 1, Column_Mask);
		}

		grid.SetRowCount(1 + row_count + 1);

		grid.Refresh();
	}

	BOOL OnItemEdit( int iRow, int iColumn )
	{
		switch (iColumn)
		{
		case Column_Mask:
			if (iRow < grid.GetRowCount() - 1)
			{
				LPARAM idx = grid.GetItemData(iRow, iColumn);
				CString text = grid.GetItemText(iRow, iColumn);

				m_view_settings->UpdateCellMask(idx, text);
				UpdateGridRows();
			}
			else
			{
				CString text = grid.GetItemText(iRow, iColumn);

				m_view_settings->AddCellMask(text);
				UpdateGridRows();
			}
			return TRUE;

		case Column_Formula:
			{

			}
		}

		return FALSE;
	}

	BOOL OnItemDelete( int iRow, int iColumn )
	{
		switch (iColumn)
		{
			case Column_Mask:
				if (iRow < grid.GetRowCount() - 1)
				{
					LPARAM idx = grid.GetItemData(iRow, iColumn);

					m_view_settings->RemoveCellMask(idx);
					UpdateGridRows();
				}
				return TRUE;
		}

		return FALSE;
	}

	bool BuildMenu(CMenu& context_menu, CPoint pt)
	{
		grid.ScreenToClient(&pt);
		CCellID cell = grid.GetCellFromPt(pt);

		if (cell.row == 0)
		{
			context_menu.CreatePopupMenu();

			for (int i = 0; i < Column_Count; ++i)
				context_menu.AppendMenu(0, ID_ShowColumn + i, GetColumnName(i));
		}

		return true;
	}

	bool IsColumnVisible(int column)
	{
		return !m_view_settings->IsColumnHidden(column);
	}

	void ToggleColumn(int column)
	{
		bool hidden = m_view_settings->IsColumnHidden(column);

		if (hidden)
		{
			m_view_settings->SetColumnHidden(column, false);
			grid.SetColumnWidth(column, m_view_settings->GetColumnWidth(column));
		}
		else
		{
			m_view_settings->SetColumnHidden(column, true);
			m_view_settings->SetColumnWidth(column, grid.GetColumnWidth(column));
			grid.SetColumnWidth(column, 0);
		}

		grid.Refresh();
	}

	IVWindowPtr	visio_window;
	IVWindowPtr	this_window;

	CGridCtrl	grid;
};


void CVisioFrameWnd::OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	m_impl->OnItemEdit(nmgv->iRow, nmgv->iColumn);
}

void CVisioFrameWnd::OnDeleteItem(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	m_impl->OnItemDelete(nmgv->iRow, nmgv->iColumn);
}

BOOL CVisioFrameWnd::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	return TRUE;
}

/**------------------------------------------------------------------------------
	Creates and initializes a new Visio (top-level) window, and subclasses the 
	"client area" of that parent Visio window (this "client area" is actually 
	a special window provided by visio for sub-classing).
-------------------------------------------------------------------------------*/

void CVisioFrameWnd::Create(IVWindowPtr window)
{
	m_impl->visio_window = window;

	HWND hwnd_parent = 
		::GetParent(GetVisioWindowHandle(window));

	CWnd* parent_window = 
		CWnd::FromHandle(hwnd_parent);

	CRect parent_rect;
	parent_window->GetClientRect(&parent_rect);

	// Construct Visio window. Make this window size a half of Visio's size
	m_impl->this_window = window->GetWindows()->Add(
		bstr_t(L"Shape Sheet Watch"), 
		static_cast<long>(visWSVisible | visWSAnchorRight | visWSAnchorTop), 
		static_cast<long>(visAnchorBarAddon), 
		static_cast<long>(parent_rect.Width()), 
		static_cast<long>(parent_rect.Height()), 
		static_cast<long>(parent_rect.Width() / 4), 
		static_cast<long>(parent_rect.Height() / 2), 
		vtMissing, vtMissing, vtMissing);

	HWND client = GetVisioWindowHandle(m_impl->this_window);
	SubclassWindow(client);

	CRect rect;
	GetClientRect(rect);

	m_impl->grid.Create(rect, this, 1, WS_CHILD|WS_VISIBLE|WS_BORDER);

	m_impl->Create(window);

	m_impl->OnSelectionChanged(window);
}

void CVisioFrameWnd::Destroy()
{
	GetParent()->SendMessage(WM_CLOSE);
}

/**-----------------------------------------------------------------------------
	Called when window size is changed.
------------------------------------------------------------------------------*/

void CVisioFrameWnd::OnSize(UINT nType, int cx, int cy)
{
	m_impl->grid.MoveWindow(0, 0, cx, cy);
}

/**-----------------------------------------------------------------------------
	All background is filled with child frame - nothing to erase.
------------------------------------------------------------------------------*/

BOOL CVisioFrameWnd::OnEraseBkgnd(CDC* pDC)
{
	return TRUE;
}

/**-----------------------------------------------------------------------------
	Called when window is about destroying - save window position.
------------------------------------------------------------------------------*/

void CVisioFrameWnd::OnDestroy()
{
	m_impl->Destroy();

	CWnd::OnDestroy();
}

void CVisioFrameWnd::PostNcDestroy()
{
	delete this;
}

CVisioFrameWnd::CVisioFrameWnd()
{
	m_impl = new Impl();
}

CVisioFrameWnd::~CVisioFrameWnd()
{
	delete m_impl;
}

void CVisioFrameWnd::OnContextMenu(CWnd* pWnd, CPoint point)
{
	CMenu context_menu;

	if (m_impl->BuildMenu(context_menu, point))
		context_menu.TrackPopupMenu(0, point.x, point.y, this);
}


void CVisioFrameWnd::OnShowColumn(UINT cmd_id)
{
	m_impl->ToggleColumn(cmd_id - ID_ShowColumn);
}

void CVisioFrameWnd::OnUpdateShowColumn(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(m_impl->IsColumnVisible(pCmdUI->m_nID - ID_ShowColumn));
}

void CVisioFrameWnd::OnInitMenuPopup(CMenu* pPopupMenu, UINT nIndex, BOOL bSysMenu) 
{ 
	if (!bSysMenu) 
	{
		CCmdUI state; 

		state.m_pMenu = pPopupMenu; 
		state.m_nIndexMax = pPopupMenu->GetMenuItemCount(); 

		for (state.m_nIndex = 0; state.m_nIndex < state.m_nIndexMax; state.m_nIndex++) 
		{ 
			state.m_nID = pPopupMenu->GetMenuItemID(state.m_nIndex); 
			state.m_pSubMenu = NULL; 
			state.DoUpdate(this, state.m_nID < 0xF000); 
		}
	} 
} 