
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
		, m_had_shapes(false)
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
	std::vector<VisioEventPtr> m_shape_events;

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

		CaptureSelectionChange(window);

		theApp.AddView(this);

		theApp.UpdateViews(0);
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
		Visio event: selection deleted
	-------------------------------------------------------------------------*/

	void OnSelectionDelete(IVSelectionPtr selection)
	{
		if (theApp.GetViewSettings()->IsFilterPinOn())
		{
			selection->PutIterationMode(visSelModeSkipSuper);
			for (long i = 1; i <= selection->Count; ++i)
			{
				CComPtr<IVShape> shape = selection->Item[i];
				m_shapes.erase(shape->ID);
			}
		}
		else
		{
			m_shapes.clear();
		}

		HookToSelectionEvents();
	}

	/**------------------------------------------------------------------------
		Visio event: cell changed
	-------------------------------------------------------------------------*/

	void OnCellChanged(IVCellPtr cell)
	{
		CString val = cell->Name;
		theApp.UpdateViews(UpdateOption_Hilight|UpdateOption_UseKey);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	CHTMLayoutCtrl* GetHtmlControl() const
	{
		return DYNAMIC_DOWNCAST(CHTMLayoutCtrl, m_this->GetParent());
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	CString GetSelectionCaption()
	{
		size_t count = m_shapes.size();

		if (count == 0)
			return L" ";

		if (count == 1)
		{
			CComPtr<IVShape> shape = m_shapes.begin()->second;
			return static_cast<LPCWSTR>(shape->Name);
		}

		return L"<Multiple selection>";
	}

	/**------------------------------------------------------------------------
		Visio event: selection changed
	-------------------------------------------------------------------------*/

	void HookToSelectionEvents()
	{
		m_shape_events.clear();

		for (Shapes::const_iterator it = m_shapes.begin(); it != m_shapes.end(); ++it)
		{
			CComPtr<IVShape> shape = it->second;

			IVEventListPtr event_list = shape->GetEventList();

			VisioEventPtr evt_cell(new CVisioEvent());
			evt_cell->Advise(event_list, (visEvtMod|visEvtCell), this);
			m_shape_events.push_back(evt_cell);

			VisioEventPtr evt_formula(new CVisioEvent());
			evt_formula->Advise(event_list, (visEvtMod|visEvtFormula), this);
			m_shape_events.push_back(evt_formula);
		}
	}

	void CaptureSelectionChange(IVWindowPtr window)
	{
		if (!theApp.GetViewSettings()->IsFilterPinOn())
		{
			m_shapes.clear();

			IVSelectionPtr selection = window->Selection;
			selection->PutIterationMode(visSelModeSkipSuper);

			for (long i = 1; i <= selection->Count; ++i)
			{
				CComPtr<IVShape> shape = selection->Item[i];
				m_shapes.insert( std::make_pair(shape->ID, shape) );
			}

			HookToSelectionEvents();
		}
	}

	void OnSelectionChanged(IVWindowPtr window)
	{
		CaptureSelectionChange(window);
		theApp.UpdateViews(UpdateOption_UseKey|UpdateOption_Hilight);
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	struct CellInfo
	{
		CellInfo()
			: data(0), image(-1), bk_color(CLR_DEFAULT), fg_color(CLR_DEFAULT)
		{
		}

		CString text;
		LPARAM	data;
		int		image;

		std::set<CString> values;

		COLORREF bk_color;
		COLORREF fg_color;
	};

	struct RowInfo
	{
		RowInfo()
			: cells(Column_Count), status(0)
		{
		}

		enum Status
		{
			Status_Local	= 1,
			Status_Updated	= 2,
		};

		int	status;
		std::vector<CellInfo> cells;
	};

	typedef std::map<shapesheet::SRC, RowInfo> RowInfos;

	/**------------------------------------------------------------------------
		Set "read-only" column text
	-------------------------------------------------------------------------*/

	void UpdateCellColors(CellInfo& cell_info, int col) const
	{
		switch (col)
		{
		case Column_Mask:
			cell_info.bk_color = COLOR_MASK_BK; 
			cell_info.fg_color = COLOR_MASK_FG; 
			return;

		case Column_Name:
		case Column_S:
		case Column_R:
		case Column_RU:
		case Column_C:

			cell_info.bk_color = COLOR_SRC_BK; 
			cell_info.fg_color = COLOR_SRC_FG; 
			return;
		}

		if (!IsCellEditable(col))
		{
			cell_info.bk_color = COLOR_RO_BK; 
			cell_info.fg_color = COLOR_RO_FG; 
			return;
		}
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

	static shapesheet::SRC GetRowInfoSRC(const RowInfo& row_info)
	{
		shapesheet::SRC src;

		src.s = (short)row_info.cells[Column_S].data;
		src.r = (short)row_info.cells[Column_R].data;
		src.c = (short)row_info.cells[Column_C].data;

		src.s_name = row_info.cells[Column_S].text;
		src.r_name_u = row_info.cells[Column_RU].text;
		src.r_name_l = row_info.cells[Column_R].text;
		src.c_name = row_info.cells[Column_C].text;

		return src;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	struct CellKey
	{
		CellKey(CShapeSheetGridCtrl* p_this, CCellID id)
		{
			col = id.col;

			if (col >= 0)
			{
				src = GetRowSRC(p_this, id.row);
				mask = p_this->GetItemText(id.row, Column_Mask);
			}
		}

		shapesheet::SRC src;
		int col;
		CString mask;

		bool operator < (const CellKey& other) const
		{
			if (src < other.src) return true;
			if (other.src < src) return false;

			if (col < other.col) return true;
			if (other.col < col) return false;

			return mask < other.mask;
		}

		bool operator == (const CellKey& other) const
		{
			if (mask != other.mask)
				return false;

			if (src.c != other.src.c) 
				return false;

			if (src.s != other.src.s || src.s_name != other.src.s_name)
				return false;

			if (col != other.col)
				return false;

			if (shapesheet::IsNamedRowSection(src.s))
				return (src.r == other.src.r);
			else
				return (src.r_name_u == other.src.r_name_u);
		}
	};

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	bool IsUpdated(const RowInfos& rows, const RowInfo& row, int col, const CString& new_val) const
	{
		RowInfos::const_iterator found = rows.find(GetRowInfoSRC(row));
		if (found == rows.end())
			return true;

		return found->second.cells[col].text != new_val;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void UpdateValueCellInfo(
		const RowInfos& rows,
		RowInfo& result, int col, 
		const shapesheet::SRC& src, int options) const
	{
		bool have_local = false;

		int missing_count = 0;
		int valid_count = 0;

		CString new_val = L"?";

		CellInfo& cell_info = result.cells[col];

		cell_info.values.clear();
		for (Shapes::const_iterator it = m_shapes.begin(); it != m_shapes.end(); ++it)
		{
			CComPtr<IVShape> shape = it->second;

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

				CString val = GetCellValue(cell, col);

				cell_info.values.insert(val);

				if (val != new_val)
				{
					if (new_val == L"?")
						new_val = val;
					else
						new_val = L"<multiselect>";
				}
			}
			else
			{
				++missing_count;
			}
		}

		if (have_local)
		{
			result.status |= RowInfo::Status_Local;
			cell_info.fg_color = RGB(0,0,255);
		}

		if (missing_count + valid_count > 1)
			result.cells[Column_Name].image = valid_count ? (missing_count ? 1 : 2) : 0;

		if (options & UpdateOption_Hilight)
		{
			if (IsUpdated(rows, result, col, new_val))
			{
				result.status |= RowInfo::Status_Updated;
				cell_info.bk_color = RGB(255,255,192);
			}
		}

		cell_info.text = new_val;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void UpdateCellInfo(RowInfo& result_row, int col, LPCWSTR text, LPARAM data) const
	{
		CellInfo& cell_info = result_row.cells[col];

		UpdateCellColors(cell_info, col);

		if (text != NULL)
			cell_info.text = text;

		cell_info.data = data;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void UpdateGridColumns()
	{
		m_this->SetFixedRowCount(1);

		m_this->SetColumnCount(Column_Count);

		for (int i = 0; i < Column_Count; ++i)
		{
			m_this->SetItemFgColour(0, i, COLOR_TH_FG);
			m_this->SetItemBkColour(0, i, COLOR_TH_BK);
			m_this->SetItemText(0, i, GetColumnName(i));

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

		src.s_name = p_this->GetItemText(iRow, Column_S);
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

		bool m_use_key;
		CCellID m_focus_id;
		CellKey m_focus_key;

		SaveFocus(CShapeSheetGridCtrl* p_this, bool use_key)
			: m_this(p_this)
			, m_use_key(use_key)
			, m_focus_key(p_this, p_this->GetBaseCellID(m_this->GetFocusCell()))
			, m_focus_id(m_this->GetFocusCell())
		{
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
					CCellID id = m_this->GetBaseCellID(CCellID(r,c));
					CellKey key(m_this, id);

					if (key == focus_key)
						RestoreFocusUsingId(id);
				}
			}
		}

		~SaveFocus()
		{
			if (!m_focus_id.IsValid())
				return;

			if (m_use_key)
				RestoreFocusUsingKey(m_focus_key);
			else
				RestoreFocusUsingId(m_focus_id);
		}
	};

	typedef std::set<shapesheet::SRC> CellSet;
	typedef std::vector<CellSet> GroupCellInfos;

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	std::vector<RowInfos> m_rows;

	void FillRows(
		RowInfos& dst_rows, 
		const RowInfos& src_rows,
		const CString& cell_name_mask, size_t i,
		const CellSet& cell_set, int options)
	{
		for (CellSet::const_iterator it = cell_set.begin(); it != cell_set.end(); ++it)
		{
			RowInfo result_row;

			const shapesheet::SRC& src = (*it);

			UpdateCellInfo(result_row, Column_Mask, cell_name_mask, i);
			UpdateCellInfo(result_row, Column_Name, src.name, src.index);
			UpdateCellInfo(result_row, Column_S, src.s_name, src.s);
			UpdateCellInfo(result_row, Column_R, src.r_name_l, src.r);
			UpdateCellInfo(result_row, Column_RU, src.r_name_u, src.r);
			UpdateCellInfo(result_row, Column_C, src.c_name, src.c);

			if (CellExists(src))
			{
				UpdateValueCellInfo(src_rows, result_row, Column_Formula, src, options);
				UpdateValueCellInfo(src_rows, result_row, Column_FormulaU, src, options);
				UpdateValueCellInfo(src_rows, result_row, Column_Value, src, options);
				UpdateValueCellInfo(src_rows, result_row, Column_ValueU, src, options);
			}

			dst_rows.insert( std::make_pair(GetRowInfoSRC(result_row), result_row) );
		}
	}

	bool RowContainsFilterText(const RowInfo& result_row) const
	{
		CString filter_text = theApp.GetViewSettings()->GetFilterText();

		if (filter_text.IsEmpty())
			return true;

		for (int i = 0; i < Column_Count; ++i)
		{
			if (!theApp.GetViewSettings()->IsColumnVisible(i))
				continue;

			CString cell_text = result_row.cells[i].text;

			if (StrStrI(cell_text, filter_text))
				return true;
		}

		return false;
	}

	void FilterRows(RowInfos& dst_rows, const RowInfos& src_rows)
	{
		for (RowInfos::const_iterator it = src_rows.begin(); it != src_rows.end(); ++it)
		{
			const RowInfo& result_row = it->second;

			if (theApp.GetViewSettings()->IsFilterLocalOn() && (result_row.status & RowInfo::Status_Local) == 0)
				continue;

			if (theApp.GetViewSettings()->IsFilterUpdatedOn() && (result_row.status & RowInfo::Status_Updated) == 0)
				continue;

			if (!RowContainsFilterText(result_row))
				continue;

			dst_rows.insert( std::make_pair(GetRowInfoSRC(result_row), result_row) );
		}
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	void FillGridRows(const RowInfos& rows)
	{
		int row = m_this->GetRowCount() - 1;

		int m_start = row;
		int m_count = 0;
		CString m_last = L"{3E299FD6-E96E-40b5-A861-CD0222D64278}";

		int s_start = row;
		int s_count = 0;
		CString s_last = L"{FCA496EB-7858-452d-9B05-89E5319CC262}";

		int r_start = row;
		int r_count = 0;
		short r_last = -1;

		int row_count = static_cast<int>(rows.size());
		m_this->SetRowCount(1 + row + row_count);

		for (RowInfos::const_iterator it = rows.begin(); it != rows.end(); ++it)
		{
			const RowInfo& row_info = it->second;

			for (int col = 0; col < Column_Count; ++col)
			{
				const CellInfo& cell_info = row_info.cells[col];

				m_this->SetItemText(1 + row, col, cell_info.text);
				m_this->SetItemImage(1 + row, col, cell_info.image);
				m_this->SetItemBkColour(1 + row, col, cell_info.bk_color);
				m_this->SetItemFgColour(1 + row, col, cell_info.fg_color);
				m_this->SetItemData(1 + row, col, cell_info.data);
			}

			shapesheet::SRC src = GetRowInfoSRC(row_info);
			const CString& mask = row_info.cells[Column_Mask].text;

			if (r_last == src.r && s_last == src.s_name && m_last == mask)
				++r_count;
			else
			{
				if (r_count > 1)
					m_this->MergeCells(1 + r_start, Column_R, row, Column_R);

				r_last = src.r;
				r_start = row;
				r_count = 1;
			}

			if (s_last == src.s_name && m_last == mask)
				++s_count;
			else
			{
				if (s_count > 1)
					m_this->MergeCells(1 + s_start, Column_S, row, Column_S);

				s_last = src.s_name;
				s_start = row;
				s_count = 1;
			}

			if (m_last == mask)
				++m_count;
			else
			{
				if (m_count > 1)
					m_this->MergeCells(1 + m_start, Column_Mask, row, Column_Mask);

				m_last = mask;
				m_start = row;
				m_count = 1;
			}

			++row;
		}

		if (m_count > 1)
			m_this->MergeCells(1 + m_start, Column_Mask, row, Column_Mask);

		if (s_count > 1)
			m_this->MergeCells(1 + s_start, Column_S, row, Column_S);

		if (r_count > 1)
			m_this->MergeCells(1 + r_start, Column_R, row, Column_R);
	}

	bool m_had_shapes;

	void Update(int options)
	{
		if (options & UpdateOption_Pin)
			CaptureSelectionChange(m_window);

		if (options & UpdateOption_Columns)
			UpdateGridColumns();

		if (!m_had_shapes)
			options &= ~UpdateOption_Hilight;

		m_had_shapes = !m_shapes.empty();

		CHTMLayoutCtrl* html = GetHtmlControl();
		if (html)
			html->SetElementText("shape-caption", GetSelectionCaption());

		SaveFocus save(m_this, (options & UpdateOption_UseKey) != 0);

		const Strings& cell_name_masks = m_view_settings->GetCellMasks();

		m_this->SetRowCount(1);

		m_rows.resize(cell_name_masks.size());
		
		for (size_t i = 0; i < cell_name_masks.size(); ++i)
		{
			const CString& cell_name_mask = cell_name_masks[i];

			CellSet cell_set;
			GetCellNames(cell_name_mask, cell_set);

			if ((options & UpdateOption_Filter) == 0)
			{
				RowInfos dst_rows;
				FillRows(dst_rows, m_rows[i], cell_name_mask, i, cell_set, options);
				std::swap(m_rows[i], dst_rows);
			}

			RowInfos rows;
			FilterRows(rows, m_rows[i]);

			if (rows.empty())
			{
				RowInfo result_row;

				UpdateCellInfo(result_row, Column_Mask, cell_name_mask, i);
				UpdateCellInfo(result_row, Column_Name, L"", Column_Name);
				UpdateCellInfo(result_row, Column_S, L"", Column_S);
				UpdateCellInfo(result_row, Column_R, L"", Column_R);
				UpdateCellInfo(result_row, Column_RU, L"", Column_RU);
				UpdateCellInfo(result_row, Column_C, L"", Column_C);

				rows.insert( std::make_pair(GetRowInfoSRC(result_row), result_row) );
			}

			FillGridRows(rows);
		}

		m_this->SetRowCount(m_this->GetRowCount() + 1);

		m_this->SetHighlightText(theApp.GetViewSettings()->GetFilterText());
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

		Update(0);
		return TRUE;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	bool SetFormula(int iRow, int iColumn, bstr_t text)
	{
		shapesheet::SRC src = GetRowSRC(m_this, iRow);

		bool result = true;

		CHTMLayoutCtrl* html = GetHtmlControl();

		if (html)
			html->HidePopup("#error-msg");

		for (Shapes::const_iterator it = m_shapes.begin(); it != m_shapes.end(); ++it)
		{
			CComPtr<IVShape> shape = it->second;

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
				CRect rc;

				m_this->GetCellRect(iRow, iColumn, &rc);
				m_this->ClientToScreen(&rc);

				if (html)
				{
					html->ScreenToClient(&rc);
					html->ShowPopup("#error-msg", "#error-msg-text", rc.TopLeft(), FormatErrorMessage(e));
				}

				result = false;
			}
		}

		return result;
	}

	LRESULT SetCellFormula(int iRow, int iColumn, LPCTSTR text)
	{
		VisioScopeLock lock(theApp.GetVisioApp(), L"Edit Formula");

		if (iRow >= m_this->GetRowCount() - 1)
			return FALSE;

		BOOL success = SetFormula(iRow, iColumn, text);
		
		if (success)
			lock.Commit();

		if (success)
			theApp.UpdateViews(UpdateOption_Hilight);

		return success ? 1 : -1;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	bool IsCellEditable(int iColumn) const
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

	void GetCellNames(const CString& cell_name_mask, CellSet& result)
	{
		for (Shapes::const_iterator it = m_shapes.begin(); it != m_shapes.end(); ++it)
		{
			CComPtr<IVShape> shape = it->second;

			shapesheet::GetCellNames(shape, cell_name_mask, result);
		}
	}

	bool CellExists(const shapesheet::SRC& src) const
	{
		for (Shapes::const_iterator it = m_shapes.begin(); it != m_shapes.end(); ++it)
		{
			CComPtr<IVShape> shape = it->second;

			if (shapesheet::CellExists(shape, src))
				return true;
		}
		
		return false;
	}

	/**------------------------------------------------------------------------
		
	-------------------------------------------------------------------------*/

	LRESULT OnBeginItemEdit(int iRow, int iColumn, Strings* arrOptions)
	{
		if (!IsCellEditable(iColumn))
			return -1;

		switch (iColumn)
		{
		case Column_Mask:
			if (arrOptions)
			{
				CellSet srcs;
				GetCellNames(L"*", srcs);

				Strings names;
				for (CellSet::const_iterator it = srcs.begin(); it != srcs.end(); ++it)
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

				if (arrOptions->empty())
				{
					LPARAM mask = m_this->GetItemData(iRow, Column_Mask);
					if (m_this->GetItemText(iRow, iColumn) == L"<multiselect>")
					{
						RowInfos& row_infos = m_rows[mask];

						shapesheet::SRC src = GetRowSRC(m_this, iRow);

						const RowInfos::const_iterator found = row_infos.find(src);
						if (found != row_infos.end())
						{
							const std::set<CString>& values = found->second.cells[iColumn].values;
							arrOptions->insert(arrOptions->end(), values.begin(), values.end());
						}
					}
				}
			}
			return TRUE;

		default:
			return -1;
		}
	}

	LRESULT OnCancelItemEdit(int iRow, int iColumn)
	{
		CHTMLayoutCtrl* html = GetHtmlControl();

		if (html)
			html->HidePopup("#error-msg");

		return 0;
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

	LRESULT OnItemDelete( int iRow, int iColumn )
	{
		switch (iColumn)
		{
		case Column_Mask:
			if (iRow < m_this->GetRowCount() - 1)
			{
				LPARAM idx = m_this->GetItemData(iRow, iColumn);

				m_view_settings->RemoveCellMask(idx);
				Update(UpdateOption_Hilight);
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

	typedef std::map<long, CAdapt<CComPtr<IVShape> > > Shapes;
	Shapes m_shapes;

	ViewSettings* m_view_settings;

	CShapeSheetGridCtrl	*m_this;
};


/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

BEGIN_MESSAGE_MAP(CShapeSheetGridCtrl, CGridCtrl)
	ON_NOTIFY_REFLECT(GVN_BEGINLABELEDIT, OnBeginLabelEdit)
	ON_NOTIFY_REFLECT(GVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(GVN_CANCELLABELEDIT, OnCancelLabelEdit)
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

void CShapeSheetGridCtrl::OnCancelLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;
	*result = m_impl->OnCancelItemEdit(nmgv->iRow, nmgv->iColumn);
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

CCellID CShapeSheetGridCtrl::GetBaseCellID( const CCellID& id ) const
{
	CGridCellBase* cell = GetCell(id.row, id.col);

	if (cell && cell->IsMerged())
		return cell->GetMergeRange().GetTopLeft();
	else
		return id;
}
