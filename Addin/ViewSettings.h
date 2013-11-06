
enum Column 
{
	Column_Mask,
	Column_S,
	Column_R,
	Column_C,
	Column_Formula,
	Column_Value,

	Column_Name,
	Column_Count
};

struct ViewSettings
{
public:
	const Strings& GetCellMasks() const;

	void AddCellMask(LPCWSTR name);
	void UpdateCellMask(size_t idx, LPCWSTR text);
	void RemoveCellMask(size_t idx);

	void Save();
	void Load();

private:
	Strings m_cell_name_masks;
};
