
/*******************************************************************************

	@file Implementation of a Visio frame window.

*******************************************************************************/

#include "stdafx.h"
#include "VisioFrameWnd.h"
#include "lib/GridCtrl/GridCtrl.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
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

	typedef std::vector<CString> CellNames;
	CellNames cell_names;

	void AddCellName(LPCWSTR cell_name)
	{
		cell_names.push_back(cell_name);
	}

	enum Column 
	{
		Column_Name,
		Column_LocalName,
		Column_Formula,
		Column_Value,

		Column_Count
	};

	CString GetColumnName(int i)
	{
		switch (i)
		{
		case Column_Name:		return L"FullName";
		case Column_LocalName:	return L"Name";
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

		if (selection->Count == 0)
			return;

		Visio::IVShapePtr shape = selection->Item[1];

		struct CellInfo 
		{
			CString group_name;
			CString cell_name;
		};

		typedef std::vector<CellInfo> CellInfos;
		CellInfos cell_infos;

		for (CellNames::const_iterator it = cell_names.begin(); it != cell_names.end(); ++it)
		{
			for (short s = 0; s <= 255; ++s)
			{
				Visio::IVSectionPtr section;
				if (SUCCEEDED(shape->get_Section(s, &section)))
				{
					for (short r = 0; r < section->Count; ++r)
					{
						Visio::IVRowPtr row;
						if (SUCCEEDED(section->get_Row(r, &row)))
						{
							for (long c = 0; c < row->Count; ++c)
							{
								Visio::IVCellPtr cell;
								if (SUCCEEDED(row->get_CellU(variant_t(c), &cell)))
								{
									CString cell_name = cell->Name;

									if (StringIsLike(*it, cell_name))
									{
										CellInfo cell_info;
										cell_info.group_name = *it;
										cell_info.cell_name = cell_name;

										cell_infos.push_back(cell_info);
									}
								}
							}
						}
					}
				}
			}
		}

		grid.SetRowCount(1 + cell_infos.size() + 1);

		size_t group_start = 1;
		CString last_group_name;

		size_t i = 0; 
		while (i < cell_infos.size())
		{
			CString group_name = cell_infos[i].group_name;

			grid.SetItemText(1 + i, Column_Name, group_name);

			bstr_t cell_name = cell_infos[i].cell_name;

			if (shape->GetCellExistsU(cell_name, VARIANT_FALSE))
			{
				Visio::IVCellPtr cell = shape->GetCellsU(cell_name);

				grid.SetItemText(1 + i, Column_LocalName, cell->LocalName);
				grid.SetItemText(1 + i, Column_Formula, cell->FormulaU);
				grid.SetItemText(1 + i, Column_Value, cell->ResultStr[0]);
			}

			if (last_group_name != group_name && !last_group_name.IsEmpty())
			{
				if (group_start < 1 + i)
					grid.MergeCells(group_start, Column_Name, i, Column_Name);

				group_start = i + 1;
			}

			last_group_name = group_name;
			++i;
		}

		if (group_start < 1 + i)
			grid.MergeCells(group_start, Column_Name, i, Column_Name);

		grid.Refresh();
	}

	BOOL OnItemEdit( int iRow, int iColumn )
	{
		if (iRow == grid.GetRowCount() - 1 && iColumn == Column_Name)
		{
			CString text = grid.GetItemText(iRow, iColumn);
			AddCellName(text);
			UpdateGridRows();
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

BOOL CVisioFrameWnd::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	return FALSE;
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

	m_impl->grid.Create(rect, this, 1, WS_CHILD|WS_VISIBLE);

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
