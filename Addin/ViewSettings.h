
enum Column 
{
	Column_Mask,
	Column_Name,
	Column_S,
	Column_R,
	Column_RU,
	Column_C,
	Column_Formula,
	Column_FormulaU,
	Column_Value,
	Column_ValueU,

	Column_Count
};

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

	void SetColumnHidden(int column, bool visible);
	bool IsColumnHidden(int column) const;

	void Save();
	void Load();

private:
	std::vector<bool> m_column_hidden;
	std::vector<int> m_column_widths;
	Strings m_cell_name_masks;
};
