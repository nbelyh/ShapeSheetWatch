
#pragma once
#include "lib\GridCtrl\GridCtrl.h"

class CShapeSheetGrid
	: public CGridCtrl
{
public:
	CShapeSheetGrid(LPCWSTR name);
	HWND Create(HWND parent, UINT control_id);

	CString GetName() const;

private:
	CString m_name;
};