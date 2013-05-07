
#include "stdafx.h"
#include "ShapeSheetInfo.h"
#include "Utils.h"

using namespace Visio;

CellInfo User_Defined_Cells[] = {

	{ L"Value",		visSectionUser, visRowUser, visUserValue },
	{ L"Prompt",	visSectionUser, visRowUser, visUserPrompt },

};


CellInfo Custom_Properties[] = {

	{ L"Label",		visSectionProp, visRowProp , visCustPropsLabel },
	{ L"Prompt",	visSectionProp, visRowProp , visCustPropsPrompt },
	{ L"SortKey",	visSectionProp, visRowProp , visCustPropsSortKey },
	{ L"Type",		visSectionProp, visRowProp , visCustPropsType },
	{ L"Format",	visSectionProp, visRowProp , visCustPropsFormat },
	{ L"Value",		visSectionProp, visRowProp , visCustPropsValue },
	{ L"Invisible", visSectionProp, visRowProp , visCustPropsInvis },
	{ L"Verify",	visSectionProp, visRowProp , visCustPropsAsk },
	{ L"LangID",	visSectionProp, visRowProp , visCustPropsLangID },
	{ L"Calendar",	visSectionProp, visRowProp , visCustPropsCalendar },

};

BlockInfo blocks[] = {

	{	
		visSectionUser,
		false, true,
		L"User_Defined_Cells", 
		L"User-Defined Cells", 
		L"{name}",
		User_Defined_Cells, 
		countof(User_Defined_Cells), 
	},
	{ 
		visSectionProp,
		false, true,
		L"Custom_Properties", 
		L"Custom Properties", 
		L"{name}",
		Custom_Properties, 
		countof(Custom_Properties),
	},
};

size_t GetBlockCount()
{
	return countof(blocks);
}

const BlockInfo& GetBlockInfo(size_t i)
{
	return blocks[i];
}
