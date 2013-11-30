
#include "lib/GridCtrl/GridCtrl.h"

#pragma once

class CShapeSheetGridCtrl
	: public CGridCtrl
{
public:
	CShapeSheetGridCtrl();
	~CShapeSheetGridCtrl();

	BOOL Create(CWnd* parent, UINT id, IVWindowPtr window);
	void Destroy();

	CCellID GetBaseCellID(const CCellID& id) const;

protected:
	void OnBeginLabelEdit(NMHDR*nmhdr, LRESULT* result);
	void OnEndLabelEdit(NMHDR*nmhdr, LRESULT* result);
	void OnDeleteItem(NMHDR*nmhdr, LRESULT* result);
	void OnEndColumnWidthEdit(NMHDR*nmhdr, LRESULT* result);

	DECLARE_MESSAGE_MAP()
private:
	struct Impl;
	Impl* m_impl;
};
