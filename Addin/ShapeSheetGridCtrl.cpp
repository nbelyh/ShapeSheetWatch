
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

#define COLOR_TH_BK		RGB(153,180,209)
#define COLOR_TH_FG		RGB(0,0,0)

#define COLOR_SRC_BK	RGB(0xE3,0xE3,0xE3)
#define COLOR_SRC_FG	RGB(0xF0,0,0)

#define COLOR_RO_BK		RGB(0xF7,0xF7,0xF7)
#define COLOR_RO_FG		RGB(0x10,0x10,0x10)

#define COLOR_MASK_BK	RGB(0xFF,0xFF,0xFF)
#define COLOR_MASK_FG	RGB(0,0,0)

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

	CVisioEvent m_evt_sel_deleted;
	CVisioEvent	m_evt_sel_changed;

	typedef Ptr<CVisioEvent> VisioEventPtr;
	std::vector<VisioEventPtr> m_evt_cell_changed;

	IVWindowPtr m_window;

	CImageList m_images;

	void Attach(IVWindowPtr window)
	{
		m_window = window;

		m_images.Create(IDB_CHECKBOXES, 16, 16, RGB(0,0,0));
		m_this->SetImageList(&m_images);

		IVEventListPtr event_list = window->GetEventList();
		m_evt_sel_changed.Advise(event_list, visEvtCodeWinSelChange, this);

		IVEventListPtr doc_event_list = window->GetDocument()->GetEventList();
		m_evt_sel_deleted.Advise(doc_event_list, visEvtCodeBefSelDel, this);

		m_view_settings = theApp.GetViewSettings();

		UpdateGridColumns();

		OnSelectionChanged(window);

		theApp.AddView(this);
	}

	void Detach()
	{
		m_evt_sel_changed.Unadvise();
		m_evt_sel_deleted.Unadvise();

		m_this->DeleteAllItems();

		theApp.DelView(this);

		m_window = NULL;
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

		case UpdateHint_Rows:
			UpdateGridRows(true);
		}
	}

	/**------------------------------------------------------------------------
		Visio event: selection deleted
	-------------------------------------------------------------------------*/

	void OnSelectionDelete(IVSelectionPtr selection)
	{
		m_evt_cell_changed.clear();
		m_shape_ids.clear();
	}

	/**------------------------------------------------------------------------
		Visio event: cell changed
	-------------------------------------------------------------------------*/

	void OnCellChanged(IVCellPtr cell)
	{
		theApp.UpdateViews(UpdateHint_Rows);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	CHTMLayoutCtrl* GetHtmlControl() const
	{
		return DYNAMIC_DOWNCAST(CHTMLayoutCtrl, m_this->GetParent());
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	CString GetSelectionCaption(IVSelectionPtr selection)
	{
		long count = selection->Count;

		if (count == 0)
			return L"";

		if (count <= 3)
		{
			Strings names;
			for (long i = 1; i <= selection->Count; ++i)
				names.push_back(static_cast<LPCWSTR>(selection->Item[i]->Name));
			return JoinList(names, L", ");
		}

		return L"<multiple selection>";
	}

	/**------------------------------------------------------------------------
		Visio event: selection changed
	-------------------------------------------------------------------------*/

	void OnSelectionChanged(IVWindowPtr window)
	{
		m_evt_cell_changed.clear();
		m_shape_ids.clear();

		IVSelectionPtr selection = window->Selection;
		selection->PutIterationMode(visSelModeSkipSuper);

		for (long i = 1; i <= selection->Count; ++i)
		{
			IVShapePtr shape = selection->Item[i];

			m_shape_ids.insert(shape->ID);

			IVEventListPtr event_list = shape->GetEventList();

			VisioEventPtr evt(new CVisioEvent());
			evt->Advise(event_list, (visEvtMod|visEvtCell), this);
			m_evt_cell_changed.push_back(evt);
		}

		CHTMLayoutCtrl* html = GetHtmlControl();
		if (html)
			html->SetElementText("shape-caption", GetSelectionCaption(selection));

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

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

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

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void UpdateValueCellText(int row, int col, const shapesheet::SRC& src)
	{
		bool have_local = false;

		int missing_count = 0;
		int valid_count = 0;

		bstr_t common_val = L"?";

		for (std::set<long>::const_iterator it = m_shape_ids.begin(); it != m_shape_ids.end(); ++it)
		{
			IVShapePtr shape = GetShapeFromId(*it);

			if (shapesheet::CellExists(shape, src))
			{
				++valid_count;

				IVCellPtr cell = shapesheet::GetShapeCell(shape, src);

				IVCellPtr inherited_from_cell;
				if (SUCCEEDED(cell->get_InheritedFormulaSource(&inherited_from_cell)))
				{
					if (cell == inherited_from_cell)
						have_local = true;
				}

				bstr_t val = GetCellValue(cell, col);

				if (val != common_val)
				{
					if (common_val == bstr_t(L"?"))
						common_val = val;
					else
						common_val = L"<multiselect>";
				}
			}
			else
			{
				++missing_count;
			}
		}

		if (have_local)
			m_this->SetItemFgColour(row, col, RGB(0,0,255));

		if (missing_count + valid_count > 1)
		{
			m_this->SetItemImage(row, Column_Name, valid_count ? (missing_count ? 1 : 2) : 0);
		}

		m_this->SetItemText(row, col, common_val);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

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

	static shapesheet::SRC GetRowSRC(CShapeSheetGridCtrl* p_this, int iRow)
	{
		shapesheet::SRC src;

		src.s = (short)p_this->GetItemData(iRow, Column_S);
		src.r = (short)p_this->GetItemData(iRow, Column_R);
		src.c = (short)p_this->GetItemData(iRow, Column_C);

		src.c_name = p_this->GetItemText(iRow, Column_S);
		src.r_name_u = p_this->GetItemText(iRow, Column_RU);
		src.r_name_l = p_this->GetItemText(iRow, Column_R);
		src.c_name = p_this->GetItemText(iRow, Column_C);

		return src;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	struct SaveFocus
	{
		CShapeSheetGridCtrl* m_this;

		struct CellKey
		{
			shapesheet::SRC src;

			int col;
			CString mask;
		};

		CellKey GetCellKey(CCellID id) const
		{
			CellKey key;
			key.col = id.col;

			if (id.col >= 0)
			{
				key.src = GetRowSRC(m_this, id.row);
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

					if ( 
						key.src == focus_key.src &&
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

		const Strings& cell_name_masks = m_view_settings->GetCellMasks();

		typedef std::vector< std::set<shapesheet::SRC> > GroupCellInfos;
		GroupCellInfos cell_names(cell_name_masks.size());

		for (size_t i = 0; i < cell_name_masks.size(); ++i)
			GetCellNames(cell_name_masks[i], cell_names[i]);

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

				for (std::set<shapesheet::SRC>::const_iterator it = cell_names[i].begin(); it != cell_names[i].end(); ++it)
				{
					const shapesheet::SRC& src = (*it);

					UpdateCellText(row, Column_Name, src.name);
					m_this->SetItemData(row, Column_Name, src.index);

					UpdateCellText(row, Column_S, src.s_name);
					UpdateCellText(row, Column_R, src.r_name_l);
					UpdateCellText(row, Column_RU, src.r_name_u);
					UpdateCellText(row, Column_C, src.c_name);

					m_this->SetItemData(row, Column_S, src.s);
					m_this->SetItemData(row, Column_R, src.r);
					m_this->SetItemData(row, Column_C, src.c);
					m_this->SetItemData(row, Column_Mask, i);

					if (CellExists(src))
					{
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

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	bool SetFormula(int iRow, int iColumn, bstr_t text)
	{
		shapesheet::SRC src = GetRowSRC(m_this, iRow);

		bool result = true;

		for (std::set<long>::const_iterator it = m_shape_ids.begin(); it != m_shape_ids.end(); ++it)
		{
			IVShapePtr shape = GetShapeFromId(*it);

			if (!shapesheet::CellExists(shape, src))
				continue;

			IVCellPtr cell = shapesheet::GetShapeCell(shape, src);

			try
			{
				if (iColumn == Column_FormulaU)
					cell->PutFormulaForceU(text);
				else
					cell->PutFormulaForce(text);
			}
			catch (_com_error& e)
			{
				AfxMessageBox(FormatErrorMessage(e));
				result = false;
			}
		}

		return result;
	}

	BOOL SetCellFormula(int iRow, int iColumn, LPCTSTR text)
	{
		VisioScopeLock lock(theApp.GetVisioApp(), L"Edit Formula");

		if (iRow >= m_this->GetRowCount() - 1)
			return FALSE;

		BOOL success = SetFormula(iRow, iColumn, text);
		
		if (success)
			lock.Commit();

		UpdateGridRows(true);
		return success;
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

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	IVShapePtr GetShapeFromId(long id) const
	{
		return m_window->GetPageAsObj()->GetShapes()->GetItemFromID(id);
	}

	void GetCellNames(const CString& cell_name_mask, std::set<shapesheet::SRC>& result)
	{
		for (std::set<long>::const_iterator it = m_shape_ids.begin(); it != m_shape_ids.end(); ++it)
		{
			IVShapePtr shape = GetShapeFromId(*it);

			shapesheet::GetCellNames(shape, cell_name_mask, result);
		}
	}

	bool CellExists(const shapesheet::SRC& src) const
	{
		for (std::set<long>::const_iterator it = m_shape_ids.begin(); it != m_shape_ids.end(); ++it)
		{
			IVShapePtr shape = GetShapeFromId(*it);

			if (shapesheet::CellExists(shape, src))
				return true;
		}
		
		return false;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	LRESULT OnBeginItemEdit(int iRow, int iColumn, Strings* arrOptions)
	{
		if (!IsCellEditable(iRow, iColumn))
			return -1;

		switch (iColumn)
		{
		case Column_Mask:
			if (arrOptions)
			{
				std::set<shapesheet::SRC> srcs;
				GetCellNames(L"*", srcs);

				Strings names;
				for (std::set<shapesheet::SRC>::const_iterator it = srcs.begin(); it != srcs.end(); ++it)
				{
					const shapesheet::SRC& src = (*it);

					if (CellExists(src))
						arrOptions->push_back(src.name);
				}

				std::stable_sort(arrOptions->begin(), arrOptions->end());
			}
			return TRUE;

		case Column_Formula:
		case Column_FormulaU:

			if (arrOptions)
			{
				LPARAM index = m_this->GetItemData(iRow, Column_Name);

				const shapesheet::SSInfo& ss_info = shapesheet::GetSSInfo(index);

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
			return SetCellFormula(iRow, iColumn, m_this->GetItemText(iRow, iColumn));
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
			return SetCellFormula(iRow, iColumn, L"No Formula");
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

	std::set<long> m_shape_ids;
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
