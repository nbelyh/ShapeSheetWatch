
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

using namespace Visio;

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
		Column_Name,
		Column_RowName,
		Column_LocalName,
		Column_Formula,
		Column_Value,

		Column_Count
	};

	CString GetColumnName(int i)
	{
		switch (i)
		{
		case Column_Name:		return L"Mask";
		case Column_RowName:	return L"Row";
		case Column_LocalName:	return L"Cell";
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
			grid.SetItemBkColour(0, i, RGB(153,180,209));
			grid.SetItemFgColour(0, i, RGB(0x00,0x00,0x00));
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

	void UpdateGridRows()
	{
		Visio::IVSelectionPtr selection = visio_window->Selection;

		Visio::IVShapePtr shape = NULL;
		
		if (selection->Count > 0)
			shape = selection->Item[1];

		typedef std::vector<Strings> GroupCellInfos;
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
			grid.SetItemText(row, Column_Name, cell_name_masks[i]);

			size_t start_row = row;

			if (cell_names[i].empty())
			{
				grid.SetItemData(row, Column_Name, i);
				++row;
			}
			else for (size_t j = 0; j < cell_names[i].size(); ++j)
			{
				bstr_t cell_name = cell_names[i][j];

				grid.SetItemText(row, Column_LocalName, cell_name);

				if (shape->GetCellExistsU(cell_name, VARIANT_FALSE))
				{
					Visio::IVCellPtr cell = shape->GetCellsU(cell_name);

					grid.SetItemData(row, Column_Name, i);

					CComBSTR row_name;
					if (SUCCEEDED(cell->get_RowNameU(&row_name)))
						grid.SetItemText(row, Column_RowName, row_name);

					grid.SetItemText(row, Column_Formula, cell->FormulaU);
					grid.SetItemText(row, Column_Value, cell->ResultStr[0]);
				}

				++row;
			}

			if (row > start_row + 1)
				grid.MergeCells(start_row, Column_Name, row - 1, Column_Name);
		}

		grid.SetRowCount(1 + row_count + 1);

		grid.Refresh();
	}

	BOOL OnItemEdit( int iRow, int iColumn )
	{
		switch (iColumn)
		{
		case Column_Name:
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
			case Column_Name:
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


	Visio::IVWindowPtr	visio_window;
	Visio::IVWindowPtr	this_window;

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
