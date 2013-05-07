
#pragma once

#include "lib\GridCtrl\GridCtrl.h"
#include "ShapeSheetInfo.h"

class CShapeSheetGrid
	: public CGridCtrl
{
public:
	CShapeSheetGrid(LPCWSTR name, HELEMENT element);
	HWND Create(HWND parent, UINT control_id);

	CString GetName() const;
	void AutoSize();

	DECLARE_DYNAMIC(CShapeSheetGrid);
	DECLARE_MESSAGE_MAP();
	void OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result);
	void ReloadData();
private:
	CSimpleArray<CellInfo> m_cells;
	HELEMENT m_element;
	CString m_name;
};