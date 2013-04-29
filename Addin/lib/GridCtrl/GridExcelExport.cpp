#include "stdafx.h"
#include "GridExcelExport.h"

#include "Model/ExcelUtils.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

namespace utils {
	namespace grid {

	using namespace excel_utils;

	void GenerateData(IDispatchPtr worksheet, const CGridCtrl* grid)
	{
		struct CCellRangeLess
		{
			bool operator()(const CCellRange& left, const CCellRange& right) const
			{
				if (left.GetMinCol() < right.GetMinCol()) return false;
				if (left.GetMinCol() > right.GetMinCol()) return true;
				if (left.GetMinRow() < right.GetMinRow()) return false;
				if (left.GetMinRow() > right.GetMinRow()) return true;
				if (left.GetMaxCol() < right.GetMaxCol()) return false;
				if (left.GetMaxCol() > right.GetMaxCol()) return true;

				return left.GetMaxRow() < right.GetMaxRow();
			}
		};

		std::set<CCellRange, CCellRangeLess> processed_merge;

		int col_count = grid->GetColumnCount();
		int row_count = grid->GetRowCount();
		int fixed_col_count = grid->GetFixedColumnCount();
		int fixed_row_count = grid->GetFixedRowCount();

		Position pos = Position(1L,1L);

		for (int column_no = 0; column_no < col_count; ++column_no)
		{
			IDispatchPtr cells = GetPropertyDisp(worksheet, L"Cells");
			IDispatchPtr start = GetPropertyDisp(cells, L"Item", pos.row, pos.column);

			variant_t format = GetProperty(start, L"NumberFormat");

			Position cur = pos;

			if (row_count > 0)
			{
				IDispatchPtr range = GetPropertyDisp(start, L"Resize", row_count, 1);
				PutProperty(range, L"NumberFormat", format);

				bool wrap = true;
				PutProperty(range, L"WrapText", wrap);
				PutProperty(range, L"ColumnWidth", wrap ? 30L : 10L);
			}

			for (int row_no = 0; row_no < row_count; ++row_no)
			{
				CGridCellBase* cell = grid->GetCell(row_no, column_no);
				if (cell->IsMerged() && processed_merge.find(cell->GetMergeRange()) == processed_merge.end())
				{
					CCellRange m_range = cell->GetMergeRange();
					MergeCells(worksheet
						, Position(m_range.GetMinRow() + 1, m_range.GetMinCol() + 1)
						, Position(m_range.GetMaxRow() + 1, m_range.GetMaxCol() + 1));
					processed_merge.insert(m_range);
				}

				PutCellValue(worksheet, cur, variant_t(cell->GetText()));
				++cur.row;
			}

			++pos.column;

			if (row_count > 0)
			{
				IDispatchPtr range = GetPropertyDisp(start, L"Resize", row_count, 1);
				IDispatchPtr columns = GetProperty(range, L"Columns");
				Invoke(columns, L"AutoFit");
			}
		}
	}

	void OpenGridInExcel( CGridCtrl* grid )
	{
		if (grid == NULL)
			return;

		IDispatchPtr app;
		if (FAILED(app.GetActiveObject(L"Excel.Application")))
		{
			HRESULT hr = app.CreateInstance(L"Excel.Application");
			if (FAILED(hr))
				ThrowComError(hr, FormatString(IDS_ERROR_CREATE_INSTANCE, _T("Excel.Application")));
		}

		IDispatchPtr workbooks = GetPropertyDisp(app, L"Workbooks");

		Invoke(workbooks, L"Add");

		IDispatchPtr worksheet = GetPropertyDisp(app, L"ActiveSheet");

		GenerateData(worksheet, grid);

		PutProperty(app, L"Visible", true); 
		PutProperty(app, L"UserControl", true);
	}
}}