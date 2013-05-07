
#pragma once

struct CellInfo
{
	LPCWSTR cell_name;
	short section;
	short row;
	short cell;
};

struct BlockInfo
{
	short section;
	bool is_row_multiple;
	bool is_section_multiple;

	LPCWSTR block_name;
	LPCWSTR block_title;
	LPCWSTR row_name_template;

	const CellInfo* cell_infos;
	size_t cell_infos_count;
};

size_t GetBlockCount();
const BlockInfo& GetBlockInfo(size_t );
