
/*******************************************************************************

	@file Implementation of a Visio frame window.

*******************************************************************************/

#include "stdafx.h"
#include "Addin.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
#include "ShapeSheet.h"
#include "ShapeSheetGridCtrl.h"

/**-----------------------------------------------------------------------------
	Message map
------------------------------------------------------------------------------*/

#define COLOR_TH_BK	RGB(153,180,209)
#define COLOR_TH_FG	RGB(0,0,0)

#define COLOR_SRC_BK	RGB(0xE3,0xE3,0xE3)
#define COLOR_SRC_FG	RGB(0xF0,0,0)

struct CShapeSheetGridCtrl::Impl 
	: public VEventHandler
	, public IView
{
	Impl(CShapeSheetGridCtrl* p_this)
		: m_this(p_this)
	{
	}

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

	void UpdateGridColumns()
	{
		m_this->SetFixedRowCount(1);

		m_this->SetColumnCount(Column_Count);

		for (int i = 0; i < Column_Count; ++i)
		{
			m_this->SetItemBkColour(0, i, COLOR_TH_BK);
			m_this->SetItemFgColour(0, i, COLOR_TH_FG);

			m_this->SetItemText(0, i, GetColumnName(i));

			if (m_view_settings->IsColumnVisible(i))
				m_this->SetColumnWidth(i, m_view_settings->GetColumnWidth(i));
			else
				m_this->SetColumnWidth(i, 0);
		}
	}

	CVisioEvent m_evt_shape_deleted;
	CVisioEvent m_evt_cell_changed;
	CVisioEvent	m_evt_sel_changed;

	void Attach(IVWindowPtr window)
	{
		IVEventListPtr event_list = window->GetEventList();
		m_evt_sel_changed.Advise(event_list, visEvtCodeWinSelChange, this);

		IVEventListPtr doc_event_list = window->GetDocument()->GetEventList();
		m_evt_shape_deleted.Advise(doc_event_list, visEvtCodeBefSelDel, this);

		m_view_settings = theApp.GetViewSettings();

		UpdateGridColumns();

		OnSelectionChanged(window);

		theApp.AddView(this);
	}

	void Detach()
	{
		m_evt_sel_changed.Unadvise();
		m_evt_shape_deleted.Unadvise();

		m_this->DeleteAllItems();

		theApp.DelView(this);
	}

	void Update(int hint)
	{
		switch (hint)
		{
		case UpdateHint_Columns:
			UpdateGridColumns();
			UpdateGridRows();
			break;
		}
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

	void SetHeadColumn(int row, int col, const CString& text)
	{
		m_this->SetItemBkColour(row, col, COLOR_SRC_BK);
		m_this->SetItemFgColour(row, col, COLOR_SRC_FG);
		m_this->SetItemText(row, col, text);
	}

	ViewSettings* m_view_settings;

	void UpdateGridRows()
	{
		int visio_version = GetVisioVersion(theApp.GetVisioApp());

		const Strings& cell_name_masks = m_view_settings->GetCellMasks();

		typedef std::vector< std::vector<SRC> > GroupCellInfos;
		GroupCellInfos cell_names(cell_name_masks.size());

		if (m_shape != NULL)
		{
			for (size_t i = 0; i < cell_name_masks.size(); ++i)
				GetCellNames(m_shape, cell_name_masks[i], cell_names[i]);
		}

		m_this->SetRowCount(1);

		int row_count = 0;
		for (int i = 0; i < int(cell_names.size()); ++i)
		{
			int cell_names_count = 
				int(cell_names[i].size());

			row_count += cell_names_count > 0 ? cell_names_count : 1;
		}

		m_this->SetRowCount(1 + row_count);

		int row = 1;
		for (int i = 0; i < int(cell_name_masks.size()); ++i)
		{
			SetHeadColumn(row, Column_Mask, cell_name_masks[i]);

			int m_row = row;

			if (cell_names[i].empty())
			{
				m_this->SetItemData(row, Column_Mask, i);

				SetHeadColumn(row, Column_S, L"");
				SetHeadColumn(row, Column_R, L"");
				SetHeadColumn(row, Column_C, L"");

				++row;
			}
			else 
			{
				int s_start = row;
				int s_count = 0;
				short s_last = -1;

				int r_start = row;
				int r_count = 0;
				short r_last = -1;

				for (int j = 0; j < int(cell_names[i].size()); ++j)
				{
					SRC& src = cell_names[i][j];

					SetHeadColumn(row, Column_Name, src.name);

					SetHeadColumn(row, Column_S, src.s_name);
					SetHeadColumn(row, Column_R, src.r_name_l);
					SetHeadColumn(row, Column_RU, src.r_name_u);
					SetHeadColumn(row, Column_C, src.c_name);

					if (m_shape->GetCellsSRCExists(src.s, src.r, src.c, VARIANT_FALSE))
					{
						IVCellPtr cell = m_shape->GetCellsSRC(src.s, src.r, src.c);

						m_this->SetItemData(row, Column_Mask, i);

						m_this->SetItemText(row, Column_Formula, cell->Formula);
						m_this->SetItemText(row, Column_FormulaU, cell->Formula);

						m_this->SetItemText(row, Column_Value, cell->ResultStr[-1L]);

						if (visio_version > 11)
							m_this->SetItemText(row, Column_ValueU, cell->ResultStrU[-1L]);
					}

					if (s_last == src.s)
						++s_count;
					else
					{
						if (s_count > 1)
							m_this->MergeCells(s_start, Column_S, row - 1, Column_S);

						s_last = src.s;
						s_start = row;
						s_count = 1;
					}

					if (r_last == src.r)
						++r_count;
					else
					{
						if (r_count > 1)
							m_this->MergeCells(r_start, Column_R, row - 1, Column_R);

						r_last = src.r;
						r_start = row;
						r_count = 1;
					}

					++row;
				}

				if (s_count > 1)
					m_this->MergeCells(s_start, Column_S, row - 1, Column_S);

				if (r_count > 1)
					m_this->MergeCells(r_start, Column_R, row - 1, Column_R);
			}

			if (row > m_row + 1)
				m_this->MergeCells(m_row, Column_Mask, row - 1, Column_Mask);
		}

		m_this->SetRowCount(1 + row_count + 1);

		m_this->Refresh();
	}

	BOOL OnItemEdit( int iRow, int iColumn )
	{
		switch (iColumn)
		{
		case Column_Mask:
			if (iRow < m_this->GetRowCount() - 1)
			{
				LPARAM idx = m_this->GetItemData(iRow, iColumn);
				CString text = m_this->GetItemText(iRow, iColumn);

				m_view_settings->UpdateCellMask(idx, text);
				UpdateGridRows();
			}
			else
			{
				CString text = m_this->GetItemText(iRow, iColumn);

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
			if (iRow < m_this->GetRowCount() - 1)
			{
				LPARAM idx = m_this->GetItemData(iRow, iColumn);

				m_view_settings->RemoveCellMask(idx);
				UpdateGridRows();
			}
			return TRUE;
		}

		return FALSE;
	}

	void OnEndColumnWidthEdit( int iColumn )
	{
		m_view_settings->SetColumnWidth(iColumn, m_this->GetColumnWidth(iColumn));
	}

	IVWindowPtr	visio_window;
	IVWindowPtr	this_window;

	CShapeSheetGridCtrl	*m_this;
};

BEGIN_MESSAGE_MAP(CShapeSheetGridCtrl, CGridCtrl)
	ON_NOTIFY_REFLECT(GVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(GVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(GVN_ENDCOLUMWIDTHEDIT, OnEndColumnWidthEdit)
END_MESSAGE_MAP()

void CShapeSheetGridCtrl::OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	m_impl->OnItemEdit(nmgv->iRow, nmgv->iColumn);
}

void CShapeSheetGridCtrl::OnEndColumnWidthEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	m_impl->OnEndColumnWidthEdit(nmgv->iColumn);
}

void CShapeSheetGridCtrl::OnDeleteItem(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	m_impl->OnItemDelete(nmgv->iRow, nmgv->iColumn);
}

CShapeSheetGridCtrl::CShapeSheetGridCtrl()
{
	m_impl = new Impl(this);
}

CShapeSheetGridCtrl::~CShapeSheetGridCtrl()
{
	delete m_impl;
}

BOOL CShapeSheetGridCtrl::Create(CWnd* parent, UINT id, IVWindowPtr window)
{
	if (!CGridCtrl::Create(CRect(0,0,0,0), parent, id, WS_VISIBLE|WS_CHILD))
		return FALSE;

	m_impl->Attach(window);
	return TRUE;
}

void CShapeSheetGridCtrl::Destroy()
{
	m_impl->Detach();
}
