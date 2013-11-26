// InPlaceList.cpp : implementation file
//
// Written by Chris Maunder <cmaunder@mail.com>
// Copyright (c) 1998-2000. All Rights Reserved.
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
// 6 Aug 1998 - Added CComboEdit to subclass the edit control - code provided by 
//              Roelf Werkman <rdw@inn.nl>. Added nID to the constructor param list.
// 29 Nov 1998 - bug fix in onkeydown (Markus Irtenkauf)
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "InPlaceList.h"

#include "GridCtrl.h"
#include "lib/Utils.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


/////////////////////////////////////////////////////////////////////////////
// CComboEdit

CComboEdit::CComboEdit()
{
}

CComboEdit::~CComboEdit()
{
}

// Stoopid win95 accelerator key problem workaround - Matt Weagle.
BOOL CComboEdit::PreTranslateMessage(MSG* pMsg) 
{
	// Make sure that the keystrokes continue to the appropriate handlers
	if (pMsg->message == WM_KEYDOWN || pMsg->message == WM_KEYUP)
	{
		::TranslateMessage(pMsg);
		::DispatchMessage(pMsg);
		return TRUE;
	}	

	// Catch the Alt key so we don't choke if focus is going to an owner drawn button
	if (pMsg->message == WM_SYSCHAR)
		return TRUE;

	return CEdit::PreTranslateMessage(pMsg);
}

BEGIN_MESSAGE_MAP(CComboEdit, CEdit)
	//{{AFX_MSG_MAP(CComboEdit)
	ON_WM_KILLFOCUS()
	ON_WM_KEYDOWN()
	ON_WM_CHAR()
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CComboEdit message handlers

void CComboEdit::OnKillFocus(CWnd* pNewWnd) 
{
	CEdit::OnKillFocus(pNewWnd);
	
    CInPlaceList* pOwner = (CInPlaceList*) GetOwner();  // This MUST be a CInPlaceList
    if (pOwner)
        pOwner->EndEdit();	
}

void CComboEdit::OnChar(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (nChar == VK_RETURN || nChar == VK_ESCAPE)
		return;

	return CEdit::OnChar(nChar, nRepCnt, nFlags);
}

void CComboEdit::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if ((nChar == VK_PRIOR || nChar == VK_NEXT ||
		 nChar == VK_DOWN  || nChar == VK_UP   ) &&
		(GetKeyState(VK_CONTROL) < 0 && GetDlgCtrlID() == IDC_COMBOEDIT))
    {
        CWnd* pOwner = GetOwner();
        if (pOwner)
            pOwner->SendMessage(WM_KEYDOWN, nChar, nRepCnt+ (((DWORD)nFlags)<<16));
        return;
    }

	if (nChar == VK_ESCAPE) 
	{
		CWnd* pOwner = GetOwner();
		if (pOwner)
			pOwner->SendMessage(WM_KEYDOWN, nChar, nRepCnt + (((DWORD)nFlags)<<16));
		return;
	}

	if (nChar == VK_TAB || nChar == VK_RETURN || nChar == VK_ESCAPE)
	{
		CWnd* pOwner = GetOwner();
		if (pOwner)
			pOwner->SendMessage(WM_KEYDOWN, nChar, nRepCnt + (((DWORD)nFlags)<<16));
		return;
	}

	CEdit::OnKeyDown(nChar, nRepCnt, nFlags);
}



/////////////////////////////////////////////////////////////////////////////
// CComboBoxExtList

BEGIN_MESSAGE_MAP(CComboBoxExtList, CListBox)
	ON_MESSAGE(LB_FINDSTRING, CComboBoxExtList::OnLbFindString)
END_MESSAGE_MAP()

LRESULT CComboBoxExtList::OnLbFindString(WPARAM wParam, LPARAM lParam)
{
	return SendMessage(LB_FINDSTRINGEXACT, wParam, lParam);
}


/////////////////////////////////////////////////////////////////////////////
// CInPlaceList

CInPlaceList::CInPlaceList(CWnd* pParent, CRect& rect, DWORD dwStyle, UINT nID,
                           int nRow, int nColumn, 
						   Strings& Items,
						   CString sInitText, 
						   UINT nFirstChar)
{
	m_nNumLines = 2;
	m_sInitText = sInitText;
 	m_nRow		= nRow;
 	m_nCol      = nColumn;
 	m_nLastChar = 0; 
	m_bEdit = FALSE;

	// Create the combobox
 	DWORD dwComboStyle = WS_BORDER|WS_CHILD|WS_VISIBLE|WS_VSCROLL|
 					     CBS_AUTOHSCROLL | dwStyle;

	if (!Create(dwComboStyle, rect, pParent, nID)) 
		return;

	COMBOBOXINFO cbInfo;
	cbInfo.cbSize = sizeof(COMBOBOXINFO);
	GetComboBoxInfo(&cbInfo);

	m_edit.SubclassWindow(cbInfo.hwndItem);
	m_ListBox.SubclassWindow(cbInfo.hwndList);

	// Add the strings
	for (size_t i = 0; i < Items.size(); i++) 
		AddString(Items[i]);

	// Get the maximum width of the text strings
	int nMaxLength = 0;
	CClientDC dc(GetParent());
	CFont* pOldFont = dc.SelectObject(pParent->GetFont());

	for (size_t i = 0; i < Items.size(); i++) 
		nMaxLength = max(nMaxLength, dc.GetTextExtent(Items[i]).cx);

	nMaxLength += (::GetSystemMetrics(SM_CXVSCROLL) + dc.GetTextExtent(_T(" ")).cx*2);
	dc.SelectObject(pOldFont);

	// Resize the edit window and the drop down window

	SetFont(pParent->GetFont());

	SetDroppedWidth(nMaxLength);
	SetHorizontalExtent(0); // no horz scrolling

	// Set the initial text to m_sInitText
	SetWindowText(m_sInitText);		// No text selected, so restore what was there before
	ShowDropDown();

	SetFocus();

	// Added by KiteFly. When entering DBCS chars into cells the first char was being lost
	// SenMessage changed to PostMessage (John Lagerquist)
	switch (nFirstChar)
	{
	case VK_RETURN:
		break;

	default:
		PostMessage(WM_CHAR, nFirstChar);   
	}
}

CInPlaceList::~CInPlaceList()
{
}

void CInPlaceList::EndEdit()
{
    CString str;

	if (m_nLastChar != VK_ESCAPE)
		GetWindowText(str);
	else
		str = m_sInitText;
 
    // Send Notification to parent
    GV_DISPINFO dispinfo;

    dispinfo.hdr.hwndFrom = GetSafeHwnd();
    dispinfo.hdr.idFrom   = GetDlgCtrlID();
    dispinfo.hdr.code     = GVN_ENDLABELEDIT;
 
    dispinfo.item.mask    = LVIF_TEXT|LVIF_PARAM;
    dispinfo.item.row     = m_nRow;
    dispinfo.item.col     = m_nCol;
    dispinfo.item.strText = str;
    dispinfo.item.lParam  = (LPARAM) m_nLastChar; 
 
    CWnd* pOwner = GetOwner();
    if (IsWindow(pOwner->GetSafeHwnd()))
        pOwner->SendMessage(WM_NOTIFY, GetDlgCtrlID(), (LPARAM)&dispinfo );
 
    // Close this window (PostNcDestroy will delete this)
    PostMessage(WM_CLOSE, 0, 0);
}

void CInPlaceList::PostNcDestroy() 
{
	CComboBox::PostNcDestroy();

	delete this;
}

/////////////////////////////////////////////////////////////////////////////
// CInPlaceList message handlers

void CInPlaceList::OnKillFocus(CWnd* pNewWnd) 
{
	CComboBox::OnKillFocus(pNewWnd);

	COMBOBOXINFO cbInfo;
	cbInfo.cbSize = sizeof(COMBOBOXINFO);
	GetComboBoxInfo(&cbInfo);

	HWND hwnd = pNewWnd->GetSafeHwnd();
	if (cbInfo.hwndCombo == hwnd || cbInfo.hwndItem == hwnd || cbInfo.hwndList == hwnd)
		return;

    // Only end editing on change of focus if we're using the CBS_DROPDOWNLIST style
	EndEdit();
}

// If an arrow key (or associated) is pressed, then exit if
//  a) The Ctrl key was down, or
void CInPlaceList::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags) 
{
	if ((nChar == VK_PRIOR || nChar == VK_NEXT ||
		 nChar == VK_DOWN  || nChar == VK_UP   ))
	{
		m_nLastChar = nChar;
		GetParent()->SetFocus();
		return;
	}

	if (nChar == VK_ESCAPE) 
		SetWindowText(m_sInitText);	// restore previous text

	if (nChar == VK_TAB || nChar == VK_RETURN || nChar == VK_ESCAPE)
	{
		m_nLastChar = nChar;
		GetParent()->SetFocus();	// This will destroy this window
		return;
	}

	CComboBox::OnKeyDown(nChar, nRepCnt, nFlags);
}

/////////////////////////////////////////////////////////////////////////////
// CComboBoxExt

CInPlaceList::CItemData::CItemData()
	: m_dwItemData(CB_ERR)
	, m_sItem(_T(""))
	, m_bState(FALSE)
{
}

CInPlaceList::CItemData::CItemData(DWORD dwItemData, LPCTSTR lpszString, BOOL bState)
{
	m_dwItemData = dwItemData;
	m_sItem = lpszString;
	m_bState = bState;
}

CInPlaceList::CItemData::~CItemData()
{
}

BEGIN_MESSAGE_MAP(CInPlaceList, CComboBox)
	//{{AFX_MSG_MAP(CComboBoxExt)
	ON_WM_DESTROY()
	ON_CONTROL_REFLECT_EX(CBN_CLOSEUP, OnCloseup)
	ON_CONTROL_REFLECT_EX(CBN_DROPDOWN, OnDropdown)
	ON_CONTROL_REFLECT_EX(CBN_SELCHANGE, OnSelchange)
	ON_CONTROL_REFLECT_EX(CBN_EDITCHANGE, OnEditchange)
	ON_WM_KILLFOCUS()
	ON_WM_KEYDOWN()
	//}}AFX_MSG_MAP
	ON_MESSAGE(CB_SETCURSEL, OnSetCurSel)
	ON_MESSAGE(CB_ADDSTRING, OnAddString)
	ON_MESSAGE(CB_DELETESTRING, OnDeleteString)
	ON_MESSAGE(CB_RESETCONTENT, OnResetContent)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CComboBoxExt message handlers

BOOL CInPlaceList::PreTranslateMessage(MSG* pMsg) 
{
	if (WM_KEYDOWN == pMsg->message)
	{
		if (VK_RETURN == pMsg->wParam)
		{
			ShowDropDown(FALSE);
		}
	}

	return CComboBox::PreTranslateMessage(pMsg);
}

void CInPlaceList::FitDropDownToItems()
{
	if (NULL == m_ListBox.GetSafeHwnd())
		return;

	CRect rcEdit, rcDropDown, rcDropDownCli;
	GetWindowRect(rcEdit);
	m_ListBox.GetWindowRect(rcDropDown);
	m_ListBox.GetClientRect(rcDropDownCli);

	int nHeight = rcDropDown.Height() - rcDropDownCli.Height();
	const int nMaxHeight = ::GetSystemMetrics(SM_CYSCREEN) / 2;
	const int nCount = GetCount();
	for(int nIndex = 0;nIndex < nCount;++nIndex)
	{
		nHeight += GetItemHeight(nIndex);
		if(nHeight > nMaxHeight)
			break;
	}

	CRect rcDropDownNew(rcDropDown.left,rcDropDown.top,rcDropDown.right,rcDropDown.top + nHeight);
	if(rcEdit.top > rcDropDown.top && rcEdit.top != rcDropDownNew.bottom)
		rcDropDownNew.top += (rcEdit.top - rcDropDownNew.bottom);

	m_ListBox.MoveWindow(rcDropDownNew);
}

BOOL CInPlaceList::OnDropdown() 
{
	// TODO: Add your control notification handler code here

	SetCursor(LoadCursor(NULL, IDC_ARROW));

	int dx = 0;
	CSize sz(0,0);
	TEXTMETRIC tm;
	CString sLBText;
	CDC* pDC = GetDC();
	CFont* pFont = GetFont();
	// Select the listbox font, save the old font
	CFont* pOldFont = pDC->SelectObject(pFont);
	// Get the text metrics for avg char width
	pDC->GetTextMetrics(&tm);

	const int nCount = GetCount();
	for(int i = 0;i < nCount;++i)
	{
		GetLBText(i, sLBText);
		sz = pDC->GetTextExtent(sLBText);
		// Add the avg width to prevent clipping
		sz.cx += tm.tmAveCharWidth;
		if(sz.cx > dx)
			dx = sz.cx;
	}

	// Select the old font back into the DC
	pDC->SelectObject(pOldFont);
	ReleaseDC(pDC);

	// Adjust the width for the vertical scroll bar and the left and right border.
	dx += ::GetSystemMetrics(SM_CXVSCROLL) + 2 * ::GetSystemMetrics(SM_CXEDGE);
	if(GetDroppedWidth() < dx)
		SetDroppedWidth(dx);

	return Default() != 0;
}

void CInPlaceList::OnDestroy() 
{
	if(NULL != m_ListBox.GetSafeHwnd())
		m_ListBox.UnsubclassWindow();

	m_edit.UnsubclassWindow();

	CComboBox::OnDestroy();

	// TODO: Add your message handler code here

	POSITION pos = m_PtrList.GetHeadPosition();
	while(pos)
	{
		CItemData* pData = m_PtrList.GetNext(pos);
		if(NULL != pData)
			delete pData;
	}
	m_PtrList.RemoveAll();
}

LRESULT CInPlaceList::OnResetContent(WPARAM wParam, LPARAM lParam)
{
	m_sTypedText.Empty();

	POSITION pos = m_PtrList.GetHeadPosition();
	while(pos)
	{
		CItemData* pData = m_PtrList.GetNext(pos);
		if(NULL != pData)
			delete pData;
	}
	m_PtrList.RemoveAll();

	return Default();
}

BOOL CInPlaceList::OnSelchange() 
{
	// TODO: Add your control notification handler code here

	if(m_bEdit)
		return Default() != 0;

	const int nIndex = GetCurSel();
	if(nIndex >= 0 && nIndex < GetCount())
	{
		GetLBText(nIndex, m_sTypedText);
		UpdateText(m_sTypedText);
	}

	return Default() != 0;
}

BOOL CInPlaceList::OnEditchange() 
{
	// TODO: Add your control notification handler code here

	m_bEdit = TRUE;
	DWORD dwGetSel = GetEditSel();
	WORD wStart = LOWORD(dwGetSel);
	WORD wEnd	= HIWORD(dwGetSel);
	GetWindowText(m_sTypedText);
	CString sEditText(m_sTypedText);
	sEditText.MakeLower();

	CString sTemp, sFirstOccurrence;
	const BOOL bEmpty = m_sTypedText.IsEmpty();
	POSITION pos = m_PtrList.GetHeadPosition();
	while(! bEmpty && pos)
	{
		CItemData* pData = m_PtrList.GetNext(pos);
		sTemp = pData->m_sItem;
		sTemp.MakeLower();
		if (StringIsLike(sEditText + "*", sTemp))
			AddItem(pData);
		else
			DeleteItem(pData);
	}

	if(GetCount() < 1 || bEmpty)
	{
		if(GetDroppedState())
			ShowDropDown(FALSE);
		else
		{
			pos = m_PtrList.GetHeadPosition();
			while(pos)
				AddItem(m_PtrList.GetNext(pos));
		}
	}
	else
	{
		ShowDropDown();
		FitDropDownToItems();
	}

	UpdateText(m_sTypedText);
	SetEditSel(wStart,wEnd);

	m_bEdit = FALSE;

	return Default() != 0;
}

BOOL CInPlaceList::OnCloseup() 
{
	// TODO: Add your control notification handler code here

	POSITION pos = m_PtrList.GetHeadPosition();
	while(pos)
		AddItem(m_PtrList.GetNext(pos));

	return Default() != 0;
}

LRESULT CInPlaceList::OnSetCurSel(WPARAM wParam, LPARAM lParam) 
{
	// TODO: Add your control notification handler code here

	int nSelect = (int)wParam;
	if(nSelect >= 0 && GetCount() > nSelect)
		GetLBText(nSelect, m_sTypedText);

	return Default();
}

LRESULT CInPlaceList::OnAddString(WPARAM wParam, LPARAM lParam)
{
	int nIndex = (int)Default();

	POSITION pos = (POSITION)wParam;
	if(NULL == pos)
	{
		CItemData* pData = new CItemData(CB_ERR, (LPCTSTR)lParam, TRUE);
		m_PtrList.AddTail(pData);
		SetItemDataPtr(nIndex, pData);
	}

	return nIndex;
}

LRESULT CInPlaceList::OnDeleteString(WPARAM wParam, LPARAM lParam)
{
	POSITION pos = (POSITION)lParam;
	if(NULL == pos)
	{
		CItemData* pData = (CItemData*)GetItemDataPtr((int)wParam);
		if(NULL != pData)
		{
			POSITION pos = m_PtrList.Find(pData);
			if(NULL != pos)
			{
				delete pData;
				m_PtrList.RemoveAt(pos);
			}
		}
	}

	return Default();
}
// Add item in dropdown list
int CInPlaceList::AddItem(CItemData* pData)
{
	if(NULL == pData || TRUE == pData->m_bState)
		return CB_ERR;

	int nIndex = (int)SendMessage(CB_ADDSTRING, (WPARAM)m_PtrList.Find(pData), (LPARAM)(LPCTSTR)pData->m_sItem);
	if(CB_ERR == nIndex || CB_ERRSPACE == nIndex)
		return nIndex;

	pData->m_bState = TRUE;
	SetItemData(nIndex, pData->m_dwItemData);
	SetItemDataPtr(nIndex, pData);

	return nIndex;
}
// Delete item from dropdown list
int CInPlaceList::DeleteItem(CItemData* pData)
{
	int nIndex = CB_ERR;

	if(NULL == pData || FALSE == pData->m_bState)
		return nIndex;

	pData->m_bState = FALSE;
	const int nCount = GetCount();
	for(int i = 0;i < nCount;++i)
	{
		if(pData == (CItemData*)GetItemDataPtr(i))
		{
			nIndex = (int)SendMessage(CB_DELETESTRING, (WPARAM)i, (LPARAM)m_PtrList.Find(pData));
			break;
		}
	}

	return nIndex;
}

int CInPlaceList::SetItemData(int nIndex, DWORD dwItemData)
{
	if(CB_ERR == nIndex)
		return nIndex;

	CItemData* pData = (CItemData*)GetItemDataPtr(nIndex);
	if(NULL != pData)
	{
		pData->m_dwItemData = dwItemData;
		return 0;
	}

	return CB_ERR;
}

DWORD CInPlaceList::GetItemData(int nIndex) const
{
	if(CB_ERR == nIndex)
		return nIndex;

	DWORD dwItemData = CB_ERR;

	CItemData* pData = (CItemData*)GetItemDataPtr(nIndex);
	if(NULL != pData)
		dwItemData = pData->m_dwItemData;

	return dwItemData;
}

void CInPlaceList::UpdateText(const CString& text)
{
	Strings strings;
	SplitList(text, L"=", strings);

	CString val = (strings.size() == 2)
		? strings[0] : text;

	SetWindowText(val);
}

IMPLEMENT_DYNAMIC(CInPlaceList, CComboBox)
