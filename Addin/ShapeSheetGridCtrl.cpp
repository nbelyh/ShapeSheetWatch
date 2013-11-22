
/*******************************************************************************

	@file Implementation of a Visio frame window.

*******************************************************************************/

#include "stdafx.h"
#include "Addin.h"
#include "lib/Visio.h"
#include "lib/Utils.h"
#include "ShapeSheet.h"
#include "ShapeSheetGridCtrl.h"
#include "lib/HTMLayoutCtrl/HTMLayoutCtrl.h"

/**-----------------------------------------------------------------------------
	Message map
------------------------------------------------------------------------------*/

#define COLOR_TH_BK	RGB(153,180,209)
#define COLOR_TH_FG	RGB(0,0,0)

#define COLOR_SRC_BK	RGB(0xE3,0xE3,0xE3)
#define COLOR_SRC_FG	RGB(0xF0,0,0)

#define COLOR_RO_BK		RGB(0xF7,0xF7,0xF7)
#define COLOR_RO_FG		RGB(0x10,0x10,0x10)

#define COLOR_MASK_BK		RGB(0xFF,0xFF,0xFF)
#define COLOR_MASK_FG		RGB(0,0,0)

struct CShapeSheetGridCtrl::Impl 
	: public VEventHandler
	, public IView
{
	Impl(CShapeSheetGridCtrl* p_this)
		: m_this(p_this)
	{
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

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

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

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

	/**------------------------------------------------------------------------
		IView
	-------------------------------------------------------------------------*/

	void Update(int hint)
	{
		switch (hint)
		{
		case UpdateHint_Columns:
			UpdateGridColumns();
			UpdateGridRows(true);
			break;
		}
	}

	/**------------------------------------------------------------------------
		Visio event: selection deleted
	-------------------------------------------------------------------------*/

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

	/**------------------------------------------------------------------------
		Visio event: cell changed
	-------------------------------------------------------------------------*/

	void OnCellChanged(IVCellPtr cell)
	{
		UpdateGridRows(false);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	CHTMLayoutCtrl* GetHtmlControl() const
	{
		return DYNAMIC_DOWNCAST(CHTMLayoutCtrl, m_this->GetParent());
	}

	void SetShape(IVShapePtr shape)
	{
		m_evt_cell_changed.Unadvise();

		if (shape)
		{
			IVEventListPtr event_list = shape->GetEventList();
			m_evt_cell_changed.Advise(event_list, (visEvtMod|visEvtCell), this);
		}

		CHTMLayoutCtrl* html = GetHtmlControl();
		if (html)
		{
			html->SetElementText("shape-caption", shape->Name);
		}

		m_shape = shape;
	}

	/**------------------------------------------------------------------------
		Visio event: selection changed
	-------------------------------------------------------------------------*/

	void OnSelectionChanged(IVWindowPtr window)
	{
		IVSelectionPtr selection = window->Selection;

		IVShapePtr shape = (selection->Count == 1)
			? selection->Item[1] : window->PageAsObj->PageSheet;

		SetShape(shape);

		UpdateGridRows(false);
	}

	/**------------------------------------------------------------------------
		Set "read-only" column text
	-------------------------------------------------------------------------*/

	bool GetCellColors(int row, int col, COLORREF& bk, COLORREF& fg)
	{
		if (row == 0)
		{
			bk = COLOR_TH_BK; 
			fg = COLOR_TH_FG; 
			return true;
		}

		switch (col)
		{
		case Column_Mask:
			bk = COLOR_MASK_BK; 
			fg = COLOR_MASK_FG; 
			return true;

		case Column_Name:
		case Column_S:
		case Column_R:
		case Column_RU:
		case Column_C:
			bk = COLOR_SRC_BK;
			fg = COLOR_SRC_FG;
			return true;
		}

		if (!IsCellEditable(row, col))
		{
			bk = COLOR_RO_BK;
			fg = COLOR_RO_FG;
			return true;
		}

		return false;
	}

	bstr_t GetCellValue(IVCellPtr cell, int column) const
	{
		switch (column)
		{
		case Column_Formula: 
			return cell->Formula;

		case Column_FormulaU:
			return cell->FormulaU;

		case Column_Value:
			return cell->ResultStr[-1L];

		case Column_ValueU:
			if (GetVisioVersion() > 11)
				return cell->ResultStrU[-1L];

		default:
			return bstr_t();
		}
	}

	void UpdateValueCellText(int row, int col, const shapesheet::SRC& src)
	{
		IVCellPtr cell = 
			m_shape->GetCellsSRC(src.s, src.r, src.c);

		IVCellPtr inherited_from_cell = 
			cell->GetInheritedFormulaSource();

		if (cell == inherited_from_cell)
			m_this->SetItemFgColour(row, col, RGB(0,0,255));

		m_this->SetItemText(row, col, GetCellValue(cell, col));
	}

	void UpdateCellText(int row, int col, LPCWSTR text)
	{
		COLORREF bk;
		COLORREF fg;

		if (GetCellColors(row, col, bk, fg))
		{
			m_this->SetItemBkColour(row, col, bk);
			m_this->SetItemFgColour(row, col, fg);
		}

		m_this->SetItemText(row, col, text);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void UpdateGridColumns()
	{
		m_this->SetFixedRowCount(1);

		m_this->SetColumnCount(Column_Count);

		for (int i = 0; i < Column_Count; ++i)
		{
			UpdateCellText(0, i, GetColumnName(i));

			if (m_view_settings->IsColumnVisible(i))
				m_this->SetColumnWidth(i, m_view_settings->GetColumnWidth(i));
			else
				m_this->SetColumnWidth(i, 0);
		}
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	struct SaveFocus
	{
		CShapeSheetGridCtrl* m_this;

		struct CellKey
		{
			int s;
			int r;
			int c;

			int col;
			CString mask;
		};

		CellKey GetCellKey(CCellID id) const
		{
			CellKey key;
			key.col = id.col;

			if (id.col >= 0)
			{
				key.s = m_this->GetItemData(id.row, Column_S);
				key.r = m_this->GetItemData(id.row, Column_R);
				key.c = m_this->GetItemData(id.row, Column_C);
				key.mask = m_this->GetItemText(id.row, Column_Mask);
			}

			return key;
		}

		bool m_use_id;
		CCellID m_focus_id;
		CellKey m_focus_key;

		SaveFocus(CShapeSheetGridCtrl* p_this, bool use_id)
			: m_this(p_this), m_use_id(use_id)
		{
			m_focus_id = m_this->GetFocusCell();
			m_focus_key = GetCellKey(m_focus_id);
		}

		void RestoreFocusUsingId(CCellID focus_id)
		{
			m_this->SelectCells(focus_id);
			m_this->SetFocusCell(focus_id);
		}

		void RestoreFocusUsingKey(const CellKey& focus_key)
		{
			for (int r = 0; r < m_this->GetRowCount(); ++r)
			{
				for (int c = 0; c < m_this->GetColumnCount(); ++c)
				{
					CCellID id(r,c);
					CellKey key = GetCellKey(id);

					if (key.s == focus_key.s && 
						key.r == focus_key.r && 
						key.c == focus_key.c && 
						key.col == focus_key.col &&
						key.mask == focus_key.mask)
					{
						RestoreFocusUsingId(id);
					}
				}
			}
		}

		~SaveFocus()
		{
			if (!m_focus_id.IsValid())
				return;

			if (m_use_id)
				RestoreFocusUsingId(m_focus_id);
			else
				RestoreFocusUsingKey(m_focus_key);
		}
	};

	void UpdateGridRows(bool use_id)
	{
		SaveFocus save(m_this, use_id);

		using namespace shapesheet;

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
			UpdateCellText(row, Column_Mask, cell_name_masks[i]);

			int m_row = row;

			if (cell_names[i].empty())
			{
				m_this->SetItemData(row, Column_Mask, i);
				m_this->SetItemData(row, Column_Name, -1);

				UpdateCellText(row, Column_S, L"");
				UpdateCellText(row, Column_R, L"");
				UpdateCellText(row, Column_C, L"");

				m_this->SetItemData(row, Column_S, Column_S);
				m_this->SetItemData(row, Column_R, Column_R);
				m_this->SetItemData(row, Column_C, Column_C);

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

					UpdateCellText(row, Column_Name, src.name);
					m_this->SetItemData(row, Column_Name, src.index);

					UpdateCellText(row, Column_S, src.s_name);
					UpdateCellText(row, Column_R, src.r_name_l);
					UpdateCellText(row, Column_RU, src.r_name_u);
					UpdateCellText(row, Column_C, src.c_name);

					if (m_shape->GetCellsSRCExists(src.s, src.r, src.c, VARIANT_FALSE))
					{
						m_this->SetItemData(row, Column_S, src.s);
						m_this->SetItemData(row, Column_R, src.r);
						m_this->SetItemData(row, Column_C, src.c);

						m_this->SetItemData(row, Column_Mask, i);

						UpdateValueCellText(row, Column_Formula, src);
						UpdateValueCellText(row, Column_FormulaU, src);
						UpdateValueCellText(row, Column_Value, src);
						UpdateValueCellText(row, Column_ValueU, src);
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

	/**------------------------------------------------------------------------
		Grid item edited
	-------------------------------------------------------------------------*/

	BOOL OnEditMask(int iRow, int iColumn)
	{
		CString text = 
			m_this->GetItemText(iRow, iColumn);

		if (iRow < m_this->GetRowCount() - 1)
		{
			LPARAM idx = m_this->GetItemData(iRow, Column_Mask);

			m_view_settings->UpdateCellMask(idx, text);
		}
		else
		{
			m_view_settings->AddCellMask(text);
		}

		UpdateGridRows(true);
		return TRUE;
	}

	BOOL SetFormula(int iRow, int iColumn, bstr_t text)
	{
		short s = (short)m_this->GetItemData(iRow, Column_S);
		short r = (short)m_this->GetItemData(iRow, Column_R);
		short c = (short)m_this->GetItemData(iRow, Column_C);

		IVCellPtr cell = m_shape->GetCellsSRC(s, r, c);

		try
		{
			if (iColumn == Column_FormulaU)
				cell->PutFormulaForceU(text);
			else
				cell->PutFormulaForce(text);

			return TRUE;
		}
		catch (_com_error& e)
		{
			AfxMessageBox(FormatErrorMessage(e));
			return FALSE;
		}
	}

	BOOL OnEditFormula(int iRow, int iColumn)
	{
		if (iRow >= m_this->GetRowCount() - 1)
			return FALSE;

		bstr_t text = 
			m_this->GetItemText(iRow, iColumn);

		return SetFormula(iRow, iColumn, text);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	bool IsCellEditable(int iRow, int iColumn) const
	{
		switch (iColumn)
		{
		case Column_Mask:
		case Column_Formula:
		case Column_FormulaU:
			return true;

		default:
			return false;
		}
	}

	LRESULT OnBeginItemEdit(int iRow, int iColumn, Strings* arrOptions)
	{
		if (!IsCellEditable(iRow, iColumn))
			return -1;

		using namespace shapesheet;

		switch (iColumn)
		{
		case Column_Mask:
			if (m_shape && arrOptions)
			{
				std::vector<SRC> srcs;
				GetCellNames(m_shape, L"*", srcs);

				Strings names;
				for (size_t i = 0; i < srcs.size(); ++i)
				{
					if (m_shape->GetCellsSRCExists(srcs[i].s, srcs[i].r, srcs[i].c, VARIANT_FALSE))
						arrOptions->push_back(srcs[i].name);
				}

				std::stable_sort(arrOptions->begin(), arrOptions->end());
			}
			return TRUE;

		case Column_Formula:
		case Column_FormulaU:

			if (arrOptions)
			{
				int index = m_this->GetItemData(iRow, Column_Name);

				const SSInfo& ss_info = GetSSInfo(index);

				SplitList(ss_info.values, L";", *arrOptions);
			}
			return TRUE;

		default:
			return -1;
		}
	}

	LRESULT OnEndItemEdit( int iRow, int iColumn)
	{
		switch (iColumn)
		{
		case Column_Mask:
			return OnEditMask(iRow, iColumn);

		case Column_Formula:
		case Column_FormulaU:
			return OnEditFormula(iRow, iColumn);
		}

		return -1;
	}

	/**------------------------------------------------------------------------
		Delete row and remove it's mask from view settings
	-------------------------------------------------------------------------*/

	BOOL OnItemDelete( int iRow, int iColumn )
	{
		switch (iColumn)
		{
		case Column_Mask:
			if (iRow < m_this->GetRowCount() - 1)
			{
				LPARAM idx = m_this->GetItemData(iRow, iColumn);

				m_view_settings->RemoveCellMask(idx);
				UpdateGridRows(true);
			}
			return TRUE;

		case Column_Formula:
		case Column_FormulaU:
			return SetFormula(iRow, iColumn, L"No Formula");
		}

		return FALSE;
	}

	/**------------------------------------------------------------------------
		Changed column width - update settings
	-------------------------------------------------------------------------*/

	void OnEndColumnWidthEdit( int iColumn )
	{
		m_view_settings->SetColumnWidth(iColumn, m_this->GetColumnWidth(iColumn));
	}

private:

	IVShapePtr m_shape;
	ViewSettings* m_view_settings;

	CShapeSheetGridCtrl	*m_this;
};


/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

BEGIN_MESSAGE_MAP(CShapeSheetGridCtrl, CGridCtrl)
	ON_NOTIFY_REFLECT(GVN_BEGINLABELEDIT, OnBeginLabelEdit)
	ON_NOTIFY_REFLECT(GVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(GVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(GVN_ENDCOLUMWIDTHEDIT, OnEndColumnWidthEdit)
END_MESSAGE_MAP()

void CShapeSheetGridCtrl::OnBeginLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	GV_BEGINEDIT*nmgv  = (GV_BEGINEDIT*)nmhdr;
	*result = m_impl->OnBeginItemEdit(nmgv->iRow, nmgv->iColumn, nmgv->arrOptions);
}

void CShapeSheetGridCtrl::OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	*result = m_impl->OnEndItemEdit(nmgv->iRow, nmgv->iColumn);
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
