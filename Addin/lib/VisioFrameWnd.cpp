
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
END_MESSAGE_MAP()

#define COLOR_TH_BK	RGB(153,180,209)
#define COLOR_TH_FG	RGB(0,0,0)

#define COLOR_SRC_BK	RGB(0xE3,0xE3,0xE3)
#define COLOR_SRC_FG	RGB(0xF0,0,0)

struct CVisioFrameWnd::Impl : public VEventHandler
{
	/**-----------------------------------------------------------------------------
	Visio event handler callback
	------------------------------------------------------------------------------*/

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
			OnSelectionChanged();
			break;
		}

		return S_OK;

		LEAVE_METHOD()
	}

	IVShapePtr GetSelectedShape()
	{
		IVWindowPtr window;
		theApp.GetVisioApp()->get_ActiveWindow(&window);

		if (window == NULL)
			return NULL;

		IVSelectionPtr selection;
		window->get_Selection(&selection);

		if (selection == NULL)
			return NULL;

		long count;
		selection->get_Count(&count);

		if (count > 0)
		{
			IVShapePtr shape;
			if (SUCCEEDED(selection->get_Item(1, &shape)))
				return shape;
		}

		return NULL;
	}

	Strings cell_name_masks;

	void AddCellMask(LPCWSTR name)
	{
		cell_name_masks.push_back(name);
	}

	void UpdateCellGroup(size_t idx, LPCWSTR text)
	{
		cell_name_masks[idx] = text;
	}

	void RemoveCellGroup(size_t idx)
	{
		cell_name_masks.erase(cell_name_masks.begin() + idx);
	}

	enum Column 
	{
		Column_Mask,
		Column_S,
		Column_R,
		Column_C,
		Column_Formula,
		Column_Value,

		Column_Name,
		Column_Count
	};

	CString GetColumnName(int i)
	{
		switch (i)
		{
		case Column_Mask:		return L"Mask";
		case Column_Name:		return L"Name";
		case Column_S:			return L"S";
		case Column_R:			return L"R";
		case Column_C:			return L"C";
		case Column_Formula:	return L"Formula";
		case Column_Value:		return L"Value";

		default:	return L"";
		}
	}

	void BuildGrid()
	{
		grid.SetFixedRowCount(1);

		grid.SetColumnCount(Column_Count);

		for (int i = 0; i < Column_Count; ++i)
		{
			grid.SetItemBkColour(0, i, COLOR_TH_BK);
			grid.SetItemFgColour(0, i, COLOR_TH_FG);

			grid.SetItemText(0, i, GetColumnName(i));
		}
	}

	void Init()
	{
		BuildGrid();
		UpdateGridRows();
	}

	void OnSelectionChanged()
	{
		UpdateGridRows();
	}

	void SetSRCColumn(short row, short col, const CString& text)
	{
		grid.SetItemBkColour(row, col, COLOR_SRC_BK);
		grid.SetItemFgColour(row, col, COLOR_SRC_FG);
		grid.SetItemText(row, col, text);
	}

	void UpdateGridRows()
	{
		IVSelectionPtr selection = visio_window->Selection;

		IVShapePtr shape = NULL;
		
		if (selection->Count > 0)
			shape = selection->Item[1];

		typedef std::vector< std::vector<SRC> > GroupCellInfos;
		GroupCellInfos cell_names(cell_name_masks.size());

		for (size_t i = 0; i < cell_name_masks.size(); ++i)
			GetCellNames(shape, cell_name_masks[i], cell_names[i]);

		grid.SetRowCount(1);

		size_t row_count = 0;
		for (size_t i = 0; i < cell_names.size(); ++i)
			row_count += cell_names[i].size() > 0 ? cell_names[i].size() : 1;

		grid.SetRowCount(1 + row_count);

		size_t row = 1;
		for (size_t i = 0; i < cell_name_masks.size(); ++i)
		{
			grid.SetItemText(row, Column_Mask, cell_name_masks[i]);

			size_t m_row = row;

			if (cell_names[i].empty())
			{
				grid.SetItemData(row, Column_Mask, i);
				++row;
			}
			else 
			{
				size_t s_start = row;
				size_t s_count = 0;
				CString s_last = L"#";

				size_t r_start = row;
				size_t r_count = 0;
				CString r_last = L"#";

				for (size_t j = 0; j < cell_names[i].size(); ++j)
				{
					SRC& src = cell_names[i][j];

					grid.SetItemText(row, Column_Name, src.name);

					SetSRCColumn(row, Column_S, src.s_name);
					SetSRCColumn(row, Column_R, src.r_name);
					SetSRCColumn(row, Column_C, src.c_name);

					IVCellPtr cell = shape->GetCellsSRC(src.s, src.r, src.c);

					grid.SetItemData(row, Column_Mask, i);
					grid.SetItemText(row, Column_Formula, cell->FormulaU);
					grid.SetItemText(row, Column_Value, cell->ResultStr[-1]);

					if (s_last == src.s_name)
						++s_count;
					else
					{
						if (s_count > 1)
							grid.MergeCells(s_start, Column_S, row - 1, Column_S);

						s_last = src.s_name;
						s_start = row;
						s_count = 1;
					}

					if (r_last == src.r_name)
						++r_count;
					else
					{
						if (r_count > 1)
							grid.MergeCells(r_start, Column_R, row - 1, Column_R);

						r_last = src.r_name;
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
				UpdateCellGroup(idx, text);
				UpdateGridRows();
			}
			else
			{
				CString text = grid.GetItemText(iRow, iColumn);
				AddCellMask(text);
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
					RemoveCellGroup(idx);
					UpdateGridRows();
				}
				return TRUE;
		}

		return FALSE;
	}


	IVWindowPtr	visio_window;
	IVWindowPtr	this_window;

	CGridCtrl	grid;
	CVisioEvent	evt_sel_changed;
};


void CVisioFrameWnd::OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	m_impl->OnItemEdit(nmgv->iRow, nmgv->iColumn);

	/*
 	IVShapePtr shape = GetSelectedShape();
 
 	if (shape)
 	{
 		int idx = GetItemData(nmgv->iRow, nmgv->iColumn);
 		
 		// IVCellPtr cell = shape->GetCellsSRC(info.section, info.row, info.cell);
 		// bstr_t val = GetItemText(nmgv->iRow, nmgv->iColumn);
 		// cell->PutFormulaForce(val);
 
 		ReloadData();
 
 		*result = TRUE;
 	}
	*/
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

	IVEventListPtr event_list = window->GetEventList();
	m_impl->evt_sel_changed.Advise(event_list, visEvtCodeWinSelChange, m_impl);

	CRect rect;
	GetClientRect(rect);

	m_impl->grid.Create(rect, this, 1, WS_CHILD|WS_VISIBLE|WS_BORDER);

	m_impl->Init();
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
	m_impl->evt_sel_changed.Unadvise();
	m_impl->grid.DeleteAllItems();
	theApp.RegisterWindow(GetVisioWindowHandle(m_impl->visio_window), NULL);

	CWnd::OnDestroy();
}

CVisioFrameWnd::CVisioFrameWnd()
{
	m_impl = new Impl();
}

CVisioFrameWnd::~CVisioFrameWnd()
{
	delete m_impl;
}
