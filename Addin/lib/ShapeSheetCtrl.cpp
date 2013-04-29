
#include "stdafx.h"
#include "Addin.h"
#include "ShapeSheetCtrl.h"

using namespace Visio;

short GetSectionColumnCount(short section)
{
	switch (section)
	{
	case visSectionUser:	return 2;
	default:				return 0;
	}
}

CString GetSectionname(short section)
{
	switch (section)
	{
	case visSectionUser:	return L"User";
	}

	return L"";
}

CString GetColumnName(short section, short col)
{
	switch (section)
	{
	case visSectionUser:
		switch (col)
		{
		case visUserValue:	return L"Value";
		case visUserPrompt:	return L"Prompt";
		}
	}

	return L"";
}

short GetSectionId(LPCWSTR name)
{
	if (!lstrcmp(name, L"visSectionUser")) return visSectionUser;
	if (!lstrcmp(name, L"visSectionProp")) return visSectionProp;

	return visSectionInval;
}

CShapeSheetGrid::CShapeSheetGrid(LPCWSTR name)
	: m_name(name)
{
	SetRowResize(TRUE);
	SetColumnResize(TRUE);
}

HWND CShapeSheetGrid::Create(HWND parent, UINT control_id)
{
	CRect rect(0, 0, 0, 0);
	CGridCtrl::Create(rect, CWnd::FromHandle(parent), 1000, WS_VISIBLE|WS_CHILD);

	IVWindowPtr window;
	theApp.GetVisioApp()->get_ActiveWindow(&window);

	IVSelectionPtr selection;
	window->get_Selection(&selection);

	long count;
	selection->get_Count(&count);

	if (count > 0)
	{
		IVShapePtr shape;
		selection->get_Item(1, &shape);

		short s = GetSectionId(m_name);
		if (shape->GetSectionExists(s, VARIANT_FALSE))
		{
			IVSectionPtr section = shape->GetSection(section);

			short row_count = shape->GetRowCount(s);
			SetRowCount(1 + row_count);

			short col_count = GetSectionColumnCount(section);
			SetColumnCount(1 + col_count);

			SetItemBkColour(0, 0, RGB(153,180,209));
			SetItemFgColour(0, 0, RGB(0x00,0x00,0x00));
			SetItemText(0, 0, L"Name");

			SetFixedRowCount(1);

			for (int c = 0; c < col_count; ++c)
			{
				SetItemBkColour(0, 1 + c, RGB(153,180,209));
				SetItemFgColour(0, 1 + c, RGB(0x00,0x00,0x00));

				SetItemText(0, 1 + c, GetColumnName(section, c));
			}

			for (short r = 0; r < row_count; ++r)
			{
				SetItemBkColour(1 + r, 0, RGB(0xE7,0xE7,0xE7));
				SetItemFgColour(1 + r, 0, RGB(0xFF,0x00,0x00));

				IVRowPtr row = section->GetRow(1 + r);
				SetItemText(1 + r, 0, row->Name);

				for (int c = 0; c < col_count; ++c)
				{
					IVCellPtr cell = shape->GetCellsSRC(section, r, c);
					SetItemText(1 + r, 1 + c, cell->Formula);
				}
			}
		}
	}

	Refresh();

	return GetSafeHwnd();
}

CString CShapeSheetGrid::GetName() const
{
	return m_name;
}
