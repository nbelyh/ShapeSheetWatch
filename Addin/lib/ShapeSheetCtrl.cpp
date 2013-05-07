
#include "stdafx.h"
#include "Addin.h"
#include "ShapeSheetCtrl.h"
#include "Utils.h"
#include "ShapeSheetInfo.h"

using namespace Visio;

CShapeSheetGrid::CShapeSheetGrid(LPCWSTR name, HELEMENT element)
	: m_name(name), m_element(element)
{
	SetRowResize(TRUE);
	SetColumnResize(TRUE);
}

namespace {

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

}

HWND CShapeSheetGrid::Create(HWND parent, UINT control_id)
{
	ReloadData();

	CRect rect(0, 0, GetVirtualWidth(), GetVirtualHeight());
	CGridCtrl::Create(rect, CWnd::FromHandle(parent), control_id, WS_VISIBLE|WS_CHILD);

	Refresh();

	return GetSafeHwnd();
}

CString CShapeSheetGrid::GetName() const
{
	return m_name;
}

void CShapeSheetGrid::ReloadData()
{
	m_cells.RemoveAll();

	IVShapePtr shape = GetSelectedShape();

	if (!shape)
		return;

	for (size_t i = 0; i < GetBlockCount(); ++i)
	{
		const BlockInfo& block_info = GetBlockInfo(i);

		if (0 != lstrcmpi(block_info.block_name, m_name))
			continue;

		if (!shape->GetSectionExists(block_info.section, FALSE))
			continue;

		short row_count = shape->GetRowCount(block_info.section);
		SetRowCount(1 + row_count);

		short col_count = block_info.cell_infos_count;
		SetColumnCount(1 + col_count);

		SetItemBkColour(0, 0, RGB(153,180,209));
		SetItemFgColour(0, 0, RGB(0x00,0x00,0x00));
		SetItemText(0, 0, block_info.block_title);

		SetFixedRowCount(1);

		for (int c = 0; c < col_count; ++c)
		{

			CString cell_name = block_info.cell_infos[c].cell_name;

			SetItemBkColour(0, 1 + c, RGB(153,180,209));
			SetItemFgColour(0, 1 + c, RGB(0x00,0x00,0x00));
			SetItemText(0, 1 + c, cell_name);
		}

		for (short r = 0; r < row_count; ++r)
		{
			SetItemBkColour(1 + r, 0, RGB(0xE7,0xE7,0xE7));
			SetItemFgColour(1 + r, 0, RGB(0xFF,0x00,0x00));
			SetItemText(1 + r, 0, shape->GetCellsSRC(block_info.section, r, 0)->GetName());

			for (int c = 0; c < col_count; ++c)
			{
				CellInfo cell_info = block_info.cell_infos[c];
				cell_info.section = block_info.section;
				cell_info.row = r;

				IVCellPtr cell = shape->GetCellsSRC(cell_info.section, cell_info.row, cell_info.cell);
				SetItemText(1 + r, 1 + c, cell->Formula);

				COLORREF clr = shape->GetCellsSRCExists(cell_info.section, cell_info.row, cell_info.cell, VARIANT_TRUE)
					? RGB(0,0,255) : CLR_DEFAULT;

				SetItemFgColour(1 + r, 1 + c, clr);

				SetItemData(1 + r, 1 + c, m_cells.GetSize());
				m_cells.Add(cell_info);
			}
		}
	}
}

void CShapeSheetGrid::AutoSize()
{
	CGridCtrl::AutoSize();

	htmlayout::dom::element el(m_element);

	CString w;
	w.Format(L"%dpx", GetVirtualWidth());
	el.set_attribute("width", w);

	CString h;
	h.Format(L"%dpx", GetVirtualHeight());
	el.set_attribute("height", h);
}

void CShapeSheetGrid::OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result)
{
	NM_GRIDVIEW *nmgv  = (NM_GRIDVIEW *)nmhdr;

	IVShapePtr shape = GetSelectedShape();
	if (shape)
	{
		int idx = GetItemData(nmgv->iRow, nmgv->iColumn);
		CellInfo& info = m_cells[idx];
		
		IVCellPtr cell = shape->GetCellsSRC(info.section, info.row, info.cell);
		bstr_t val = GetItemText(nmgv->iRow, nmgv->iColumn);
		cell->PutFormulaForce(val);

		ReloadData();

		*result = TRUE;
	}
}

IMPLEMENT_DYNAMIC(CShapeSheetGrid, CGridCtrl)

BEGIN_MESSAGE_MAP(CShapeSheetGrid, CGridCtrl)
	ON_NOTIFY_REFLECT(GVN_ENDLABELEDIT, OnEndLabelEdit)
END_MESSAGE_MAP()