
#pragma once

struct IHTMLayoutControlManager
{
	virtual CWnd*	CreateControl(const CString& type) = 0;
	virtual bool	DestroyControl(const CString& type, CWnd* wnd) = 0;
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

	void SetElementEnabled (const char* id, bool set);

	void SetElementAttribute(const char* id, const char* attribute, LPCWSTR text);
	CString GetElementAttribute (const char* id, const char* attribute);

protected:
	virtual BOOL PreTranslateMessage(MSG* pMsg);
	DECLARE_DYNAMIC(CHTMLayoutCtrl)

private:
	struct Impl;
	Impl* m_impl;
};

#pragma comment(lib, "HTMLayout.lib")
