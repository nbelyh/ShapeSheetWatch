
/////////////////////////////////////////////////////////////////////////////
// InPlaceList.h : header file
//
// Written by Chris Maunder <cmaunder@mail.com>
// Copyright (c) 19982000. All Rights Reserved.
//
// The code contained in this file is based on the original
// CInPlaceList from http://www.codeguru.com/listview
//
// This code may be used in compiled form in any way you desire. This
// file may be redistributed unmodified by any means PROVIDING it is 
// not sold for profit without the authors written consent, and 
// providing that this notice and the authors name is included. If 
// the source code in  this file is used in any commercial application 
// then acknowledgement must be made to the author of this file 
// (in whatever form you wish).
//
// This file is provided "as is" with no expressed or implied warranty.
// The author accepts no liability for any damage/loss of business that
// this product may cause.
//
// Expect bugs!
// 
// Please use and enjoy. Please let me know of any bugs/mods/improvements 
// that you have found/implemented and I will fix/incorporate them into this
// file. 
//
/////////////////////////////////////////////////////////////////////////////

#pragma once

#define IDC_COMBOEDIT 1001

#include <afxtempl.h>

/////////////////////////////////////////////////////////////////////////////
// CComboEdit window

class CComboEdit : public CEdit
{
// Construction
public:
	CComboEdit();

// Overrides
	virtual BOOL PreTranslateMessage(MSG* pMsg);
// Implementation
public:
	virtual ~CComboEdit();

	// Generated message map functions
protected:
	//{{AFX_MSG(CComboEdit)
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
};


/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/

class CComboBoxExtList : public CListBox
{
protected:
	afx_msg LRESULT OnLbFindString(WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()
};

/**------------------------------------------------------------------------
	
-------------------------------------------------------------------------*/


class CInPlaceList : public CComboBox
{
    friend class CComboEdit;

// Construction
public:
	CInPlaceList(CWnd* pParent,         // parent
                 CRect& rect,           // dimensions & location
                 DWORD dwStyle,         // window/combobox style
                 UINT nID,              // control ID
                 int nRow, int nColumn, // row and column
				 Strings& Items,   // Items in list
                 CString sInitText,     // initial selection
				 UINT nFirstChar);      // first character to pass to control

// Attributes
public:
   CComboEdit m_edit;  // subclassed edit control

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CInPlaceList)
	protected:
	virtual void PostNcDestroy();
	//}}AFX_VIRTUAL

// Implementation
public:
	virtual ~CInPlaceList();
public:
    void EndEdit();

// Generated message map functions
protected:
	//{{AFX_MSG(CInPlaceList)
	afx_msg void OnKillFocus(CWnd* pNewWnd);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	//}}AFX_MSG

	DECLARE_DYNAMIC(CInPlaceList)
	DECLARE_MESSAGE_MAP()

private:
	int		m_nNumLines;
	CString m_sInitText;
	int		m_nRow;
	int		m_nCol;
 	UINT    m_nLastChar; 

	// Construction
public:
	virtual BOOL PreTranslateMessage(MSG* pMsg);

	// Implementation
public:
	virtual DWORD GetItemData(int nIndex) const;
	virtual int SetItemData(int nIndex, DWORD dwItemData);

protected:
	BOOL m_bEdit;
	CString m_sTypedText;
	CComboBoxExtList m_ListBox;

protected:

	class CItemData : public CObject
	{
		// Attributes
	public:
		DWORD m_dwItemData;
		CString m_sItem;
		BOOL m_bState;

		// Implementation
	public:
		CItemData();
		CItemData(DWORD dwItemData, LPCTSTR lpszString, BOOL bState);
		virtual ~CItemData();
	};

	CTypedPtrList<CPtrList, CItemData*> m_PtrList;

protected:
	virtual void FitDropDownToItems();
	virtual int AddItem(CItemData* pData);
	virtual int DeleteItem(CItemData* pData);

	// Generated message map functions
protected:
	//{{AFX_MSG(CComboBoxExt)
	afx_msg void OnDestroy();
	afx_msg BOOL OnCloseup();
	afx_msg BOOL OnDropdown();
	afx_msg BOOL OnKillfocus();
	afx_msg BOOL OnSelchange();
	afx_msg BOOL OnEditchange();
	//}}AFX_MSG
	afx_msg LRESULT OnSetCurSel(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnAddString(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnDeleteString(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT OnResetContent(WPARAM wParam, LPARAM lParam);
};
