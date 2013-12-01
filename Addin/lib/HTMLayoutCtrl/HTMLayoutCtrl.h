
#pragma once

struct IHTMLayoutControlManager
{
	virtual CWnd*	CreateControl(const CString& type) = 0;
	virtual bool	DestroyControl(const CString& type, CWnd* wnd) = 0;

	virtual	bool	OnButton(const CString& id) { return false; }
	virtual	bool	OnValueChanged(const CString& id, const CString& text) { return false; }
	virtual bool	OnCheckButton(const CString& id, bool set) { return false; }
	virtual bool	OnHyperlink(const CString& id, const CString& href) { return false; }
};

class CHTMLayoutCtrl : public CWnd
{
public:
	CHTMLayoutCtrl(IHTMLayoutControlManager* manager);
	~CHTMLayoutCtrl();

	// Generic creator
	BOOL Create(const RECT& rect, CWnd* pParentWnd, UINT nID, DWORD dwStyle);

	// Load html from memory buffer
	bool LoadHtml(const BYTE* pb, DWORD nBytes);
	bool LoadHtml(LPCWSTR text);

	// Load html from memory buffer (maybe utf-8 encoded)
	bool LoadHtmlEx(const BYTE* pb, DWORD nBytes, LPCWSTR base_path);

	// Load html from the (Unicode) string
	bool LoadHtmlEx(LPCWSTR text, LPCWSTR base_path);

	UINT GetMinHeight(UINT width) const;
	UINT GetMinWidth() const;

	// load html from file
	bool LoadHtmlFile(LPCWSTR filename);
	
	void SetElementText(const char* id, LPCWSTR text);
	CString GetElementText(const char* id) const;

	bool IstElementChecked(const char* id);
	void SetElementChecked (const char* id, bool set);

	void SetElementAttribute(const char* id, const char* attribute, LPCWSTR text);
	CString GetElementAttribute (const char* id, const char* attribute) const;

	void ShowPopup(LPCSTR popup_id, LPCSTR msg_id, CPoint pt, const CString& text);
	void HidePopup(LPCSTR popup_id);

protected:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	DECLARE_DYNAMIC(CHTMLayoutCtrl)

private:
	struct Impl;
	Impl* m_impl;
};

#pragma comment(lib, "HTMLayout.lib")
