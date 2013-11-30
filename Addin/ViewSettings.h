
enum Column 
{
	Column_Mask,
	Column_Name,
	Column_S,
	Column_RU,
	Column_R,
	Column_C,
	Column_FormulaU,
	Column_Formula,
	Column_ValueU,
	Column_Value,

	Column_Count
};

LPCSTR	GetColumnDbName(int i);
int		GetColumnFromDbName(const CString& id);

CString GetColumnName(int i);

struct ViewSettings
{
	ViewSettings();
	~ViewSettings();

public:
	const Strings& GetCellMasks() const;

	void AddCellMask(LPCWSTR name);
	void UpdateCellMask(size_t idx, LPCWSTR text);
	void RemoveCellMask(size_t idx);

	void SetColumnWidth(int column, int width);
	int GetColumnWidth(int column) const;

	void SetColumnVisible(int column, bool visible);
	bool IsColumnVisible(int column) const;

	void GetWindowRect(long& l, long& t, long& w, long& h) const;
	void SetWindowRect(long l, long t, long w, long h);

	LONG GetWindowState() const;
	void SetWindowState(LONG state);

	bool IsVisibleByDefault() const;
	void SetVisibleByDefault(bool set);

	bool IsFilterLocalOn() const;
	void SetFilterLocal(bool set);

	bool IsFilterUpdatedOn() const;
	void SetFilterUpdated(bool set);

	bool IsFilterPinOn() const;
	void SetFilterPin(bool set);

	void Save();
	void Load();

	CString GetFilterText() const;
	void SetFilterText(const CString& text);

private:
	CString m_filter_text;

	std::vector<bool> m_column_visible;
	std::vector<int> m_column_widths;

	DWORD m_filter_local;
	DWORD m_filter_udpated;
	DWORD m_filter_pin;

	DWORD m_window_top;
	DWORD m_window_left;
	DWORD m_window_width;
	DWORD m_window_height;
	DWORD m_window_state;
	DWORD m_visible_by_default;

	Strings m_cell_name_masks;
};
