
#pragma once

namespace shapesheet {

	// Information about a single ShapeSheet cell,
	// including section, row, column
	// the names of rows/cells can include placeholders: {r} for row name, {i} for for row index

	struct SSInfo
	{
		int visio_version;

		LPCWSTR s_name;	// ShapeSheet section name
		LPCWSTR r_name;	// ShapeSheet row name, can include placeholders
		LPCWSTR c_name;	// ShapeSheet column name

		short s;		// ShapeSheet section index
		short r;		// ShapeSheet row index
		short c;		// ShapeSheet column index

		LPCWSTR name;	// full name "mask" with placeholders for variable rows/cells
		LPCWSTR values;	// possible set of values for the cell (value list, semicolon separated)

		size_t index;
	};

	const SSInfo& GetSSInfo(size_t index);

	struct SRC
	{
		CString name;

		short s;
		CString s_name;

		short r;
		CString r_name_l;
		CString r_name_u;

		short c;
		CString c_name;

		size_t index;
	};

	void GetCellNames(IVShapePtr shape, const CString& cell_name_mask, std::vector<SRC>& result);

} // namespace shapesheet
