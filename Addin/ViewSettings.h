
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

	void Save();
	void Load();

private:
	std::vector<bool> m_column_visible;
	std::vector<int> m_column_widths;

	long m_window_top;
	long m_window_left;
	long m_window_width;
	long m_window_height;

	long m_window_state;

	Strings m_cell_name_masks;
};
