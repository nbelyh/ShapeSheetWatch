
#include "stdafx.h"
#include "lib/Utils.h"
#include "ViewSettings.h"

void ViewSettings::AddCellMask(LPCWSTR name)
{
	m_cell_name_masks.push_back(name);
}

void ViewSettings::UpdateCellMask(size_t idx, LPCWSTR text)
{
	m_cell_name_masks[idx] = text;
}

void ViewSettings::RemoveCellMask(size_t idx)
{
	m_cell_name_masks.erase(m_cell_name_masks.begin() + idx);
}

#define ADDIN_KEY	L"Software\\UnmanagedVisio\\ShapeSheetWatch"

void ViewSettings::Save()
{
	CRegKey key;
	if (0 == key.Create(HKEY_CURRENT_USER, ADDIN_KEY))
	{
		CString val_masks = JoinList(m_cell_name_masks, L"|");
		key.SetStringValue(L"Masks", val_masks, REG_SZ);

		Strings column_widths(Column_Count);
		Strings column_hidden(Column_Count);

		for (int i = 0; i < Column_Count; ++i)
		{
			column_widths[i] = FormatString(L"%d", m_column_widths[i]);
			column_hidden[i] = FormatString(L"%d", m_column_visible[i] != 0);
		}

		CString val_widths = JoinList(column_widths, L"|");
		key.SetStringValue(L"ColumnWidth", val_widths, REG_SZ);

		CString val_hidden = JoinList(column_hidden, L"|");
		key.SetStringValue(L"ColumnVisible", val_hidden, REG_SZ);
	}
}

void ViewSettings::Load()
{
	CRegKey key;
	if (0 == key.Open(HKEY_CURRENT_USER, ADDIN_KEY, KEY_READ))
	{
		WCHAR val[2048] = L"";
		ULONG len = 0;

		len = _countof(val);
		if (0 == key.QueryStringValue(L"Masks", val, &len))
			SplitList(val, L"|", m_cell_name_masks);

		len = _countof(val);
		Strings column_widths;
		if (0 == key.QueryStringValue(L"ColumnWidth", val, &len))
			lstrcpy(val, L"100|100|50|50|50|50|50|100|50|100");

		SplitList(val, L"|", column_widths);

		len = _countof(val);
		Strings column_hidden;
		if (0 == key.QueryStringValue(L"ColumnVisible", val, &len))
			lstrcpy(val, L"1|1|0|0|0|0|0|1|0|1");

		SplitList(val, L"|", column_hidden);

		for (size_t i = 0; i < Column_Count; ++i)
		{
			if (i < column_widths.size())
				m_column_widths[i] = StrToInt(column_widths[i]);

			if (i < column_hidden.size())
				m_column_visible[i] = StrToInt(column_hidden[i]) != 0;
		}
	}
}

const Strings& ViewSettings::GetCellMasks() const
{
	return m_cell_name_masks;
}

void ViewSettings::SetColumnWidth(int column, int width)
{
	m_column_widths[column] = width;
}

int ViewSettings::GetColumnWidth(int column) const
{
	return m_column_widths[column];
}

void ViewSettings::SetColumnVisible(int column, bool set)
{
	m_column_visible[column] = set;
}

bool ViewSettings::IsColumnVisible(int column) const
{
	return m_column_visible[column];
}

ViewSettings::ViewSettings()
	: m_column_widths(Column_Count, 50)
	, m_column_visible(Column_Count, true)
{
}

ViewSettings::~ViewSettings()
{
}

CString GetColumnName(int i)
{
	switch (i)
	{
	case Column_Mask:		return L"Mask";
	case Column_Name:		return L"Name";
	case Column_S:			return L"Section";
	case Column_R:			return L"Row";
	case Column_RU:			return L"Row (U)";
	case Column_C:			return L"Column";
	case Column_Formula:	return L"Formula";
	case Column_FormulaU:	return L"Formula (U)";
	case Column_Value:		return L"Result";
	case Column_ValueU:		return L"Result (U)";

	default:	return L"";
	}
}

LPCSTR GetColumnDbName(int i)
{
	switch (i)
	{
	case Column_Mask:		return "col_Mask";
	case Column_Name:		return "col_Name";
	case Column_S:			return "col_Section";
	case Column_R:			return "col_Row";
	case Column_RU:			return "col_RowU";
	case Column_C:			return "col_Column";
	case Column_Formula:	return "col_Formula";
	case Column_FormulaU:	return "col_FormulaU";
	case Column_Value:		return "col_Result";
	case Column_ValueU:		return "col_ResultU";

	default:	return "";
	}
}

int GetColumnFromDbName(const CString& id)
{
	for (int i = 0; i < Column_Count; ++i)
	{
		if (GetColumnDbName(i) == id)
			return i;
	}

	return -1;
}
