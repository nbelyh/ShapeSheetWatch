
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
	}
}

void ViewSettings::Load()
{
	CRegKey key;
	if (0 == key.Open(HKEY_CURRENT_USER, ADDIN_KEY, KEY_READ))
	{
		WCHAR val_masks[2048] = L"";
		ULONG len = _countof(val_masks);

		if (0 == key.QueryStringValue(L"Masks", val_masks, &len))
			SplitList(val_masks, L"|", m_cell_name_masks);
	}
}

const Strings& ViewSettings::GetCellMasks() const
{
	return m_cell_name_masks;
}

