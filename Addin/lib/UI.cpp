
#include "stdafx.h"

#include "Language.h"
#include "Addin.h"
#include "Utils.h"
#include "TextFile.h"
#include "UI.h"
#include "PictureConvert.h"

#define ADDON_NAME	L"ShapeSheetWatch"

_ATL_FUNC_INFO ClickEventInfo = { CC_STDCALL, VT_EMPTY, 2, { VT_DISPATCH, VT_BOOL|VT_BYREF } };


ClickEventRedirector::ClickEventRedirector(IUnknownPtr punk, const CString& tag) 
	: m_punk(punk), m_tag(tag)
{
	DispEventAdvise(punk);
}

ClickEventRedirector::~ClickEventRedirector()
{
	DispEventUnadvise(m_punk);
}

UINT GetCommandId(Office::_CommandBarButtonPtr button)
{
	CComBSTR parameter;
	if (SUCCEEDED(button->get_Parameter(&parameter)))
		return StrToInt(parameter);
	else
		return -1;
}

void __stdcall ClickEventRedirector::OnClick(IDispatch* pButton, VARIANT_BOOL* pCancel)
{
	try
	{
		Office::_CommandBarButtonPtr button;
		pButton->QueryInterface(__uuidof(Office::_CommandBarButton), (void**)&button);

		theApp.OnCommand(GetCommandId(button));

	}
	catch (_com_error)
	{
		// 
	}
}

struct MenuItemVisitor
{
	virtual void Visit(CMenu* menu, UINT idx) = 0;
};

void IterateMenuItems(CMenu* popup_menu, MenuItemVisitor& v)
{
	for (UINT i = 0; i < popup_menu->GetMenuItemCount(); ++i)
	{
		v.Visit(popup_menu, i);

		CMenu* sub_menu = popup_menu->GetSubMenu(i);

		if (sub_menu)
			IterateMenuItems(sub_menu, v);
	}
}

CString GetCommandTag(UINT command_id)
{
	return FormatString(L"%s_%d", ADDON_NAME, command_id);
}

void IterateMenu(UINT menu_id, MenuItemVisitor& v)
{
	CMenu menu;
	menu.LoadMenu(menu_id);
	IterateMenuItems(menu.GetSubMenu(0), v);
}

void GetTagSet(Strings* result)
{
	struct CollectAllTags : MenuItemVisitor
	{
		virtual void Visit(CMenu* menu, UINT i)
		{
			result->push_back(GetCommandTag(menu->GetMenuItemID(i)));
		}

		Strings* result;
	} v;

	v.result = result;

	IterateMenu(IDR_MENU, v);
}

/**---------------------------------------------------------------------------------
	
-----------------------------------------------------------------------------------*/

void UninstallControls(Office::_CommandBarsPtr cbs, LPCWSTR tag)
{
	Office::CommandBarControlsPtr old_controls;
	if (SUCCEEDED(cbs->FindControls(vtMissing, vtMissing, variant_t(tag), vtMissing, &old_controls)) && old_controls != NULL)
	{
		int count = 0;
		old_controls->get_Count(&count);

		for (long i = count; i > 0; --i)
		{
			Office::CommandBarControlPtr old_control;
			if (SUCCEEDED(old_controls->get_Item(variant_t(i), &old_control)))
				old_control->Delete();
		}
	}
}

void UninstallItems(Office::_CommandBarsPtr cbs)
{
	Strings tag_set;
	GetTagSet(&tag_set);

	for (Strings::const_iterator it = tag_set.begin(); it != tag_set.end(); ++it)
		UninstallControls(cbs, *it);
}


void AddinUi::InitializeItem( Office::CommandBarControlPtr item, UINT command_id)
{
	CString caption;
	caption.LoadString(command_id);
	item->put_Caption(bstr_t(caption));

	CString parameter = FormatString(L"%d", command_id);
	item->put_Parameter(bstr_t(parameter));

	// Set unique tag, so that the command is not lost
	CString tag = GetCommandTag(command_id);
	item->put_Tag(bstr_t(tag));

	CBitmap bm_picture;
	// if we are command button
	Office::MsoControlType item_type;
	if (SUCCEEDED(item->get_Type(&item_type)) && item_type == Office::msoControlButton && command_id > 0)
	{
		IPictureDispPtr picture;
		IPictureDispPtr mask;
		if (SUCCEEDED(CustomUiGetPng(MAKEINTRESOURCE(command_id), &picture, &mask)))
		{
			try
			{
				// cast to button. hopefully we will always succeed here
				Office::_CommandBarButtonPtr button = item;

				button->put_Picture(picture);
				button->put_Mask(mask);
			}
			catch (_com_error)
			{
				// There Some problems; hope to never get here.
			}
		}
	}

	UpdateItem(item);

	m_buttons.Add(ClickEventRedirectorPtr(new ClickEventRedirector(item, tag)));
}

void AddinUi::UpdateItem(Office::_CommandBarButtonPtr button)
{
	UINT id = GetCommandId(button);

	CComBSTR tag;
	button->get_Tag(&tag);

	VARIANT_BOOL new_visible = theApp.IsCommandVisible(id) ? VARIANT_TRUE : VARIANT_FALSE;

	VARIANT_BOOL old_visible = new_visible;
	button->get_Visible(&old_visible);

	if (old_visible != new_visible)
		button->put_Visible(new_visible);

	VARIANT_BOOL new_enabled = theApp.IsCommandEnabled(id) ? VARIANT_TRUE : VARIANT_FALSE;

	VARIANT_BOOL old_enabled = new_enabled;
	button->get_Enabled(&old_enabled);

	if (old_enabled != new_enabled)
		button->put_Enabled(new_enabled);

	Office::MsoButtonState new_state = 
		theApp.IsCommandChecked(id) ? Office::msoButtonDown : Office::msoButtonUp;

	Office::MsoButtonState old_state = new_state;
	button->get_State(&old_state);

	if (old_state != new_state)
		button->put_State(new_state);
}

void AddinUi::FillMenuItems( long position, Office::CommandBarControlsPtr menu_items, CMenu* popup_menu)
{
	// For each items in the menu,
	bool begin_group = false;
	for (UINT i = 0; i < popup_menu->GetMenuItemCount(); ++i)
	{
		CMenu* sub_menu = popup_menu->GetSubMenu(i);

		// set item caption
		CString item_caption;
		popup_menu->GetMenuString(i, item_caption, MF_BYPOSITION);

		// if this item is actually a separator then process next item
		if (item_caption.IsEmpty())
		{
			begin_group = true;
			continue;
		}

		// create new menu item.
		Office::CommandBarControlPtr menu_item_obj;
		menu_items->Add(
			variant_t(sub_menu ? long(Office::msoControlPopup) : long(Office::msoControlButton)), 
			vtMissing, 
			vtMissing, 
			position < 0 ? vtMissing : variant_t(position), 
			variant_t(false),
			&menu_item_obj);

		if (position > 0)
			++position;

		// obtain command id from menu
		UINT command_id = popup_menu->GetMenuItemID(i);

		// normal command; set up visio menu item
		InitializeItem(menu_item_obj, command_id);

		// if current item is first in a group, start new group
		if (begin_group)
		{
			menu_item_obj->put_BeginGroup(VARIANT_TRUE);
			begin_group = false;
		}

		// if this command has sub-menu
		if (sub_menu)
		{
			Office::CommandBarPopupPtr popup_menu_item_obj = menu_item_obj;

			Office::CommandBarControlsPtr controls;
			popup_menu_item_obj->get_Controls(&controls);

			FillMenuItems(-1, controls, sub_menu);
		}
	}
}

void AddinUi::FillMenu(long position, Office::CommandBarControlsPtr cbs, UINT menu_id)
{
	CMenu menu;
	menu.LoadMenu(menu_id);

	FillMenuItems(position, cbs, menu.GetSubMenu(0));
}

/**-----------------------------------------------------------------------------
	Installs toolbar
------------------------------------------------------------------------------*/

void AddinUi::CreateCommandBarsUI(IVApplicationPtr app)
{
	LanguageLock lock(GetAppLanguage(app));

	Office::_CommandBarsPtr cbs = app->CommandBars;

	Office::CommandBarPtr frame_toolbar;
	if (FAILED(cbs->get_Item(variant_t(ADDON_NAME), &frame_toolbar)))
	{
		cbs->Add(variant_t(ADDON_NAME), vtMissing, vtMissing, vtMissing, &frame_toolbar);

		frame_toolbar->put_Visible(variant_t(true));
		frame_toolbar->put_Context(bstr_t(L"2*"));
	}

	UninstallItems(cbs);

	Office::CommandBarControlsPtr controls;
	frame_toolbar->get_Controls(&controls);

	FillMenu(-1, controls, IDR_MENU);
}

void AddinUi::UpdateCommandBarsUI()
{
	Office::_CommandBarsPtr cbs = theApp.GetVisioApp()->GetCommandBars();

	Strings tag_set;
	GetTagSet(&tag_set);

	for (Strings::const_iterator it = tag_set.begin(); it != tag_set.end(); ++it)
	{
		CString tag = (*it);

		Office::CommandBarControlsPtr controls;
		if (SUCCEEDED(cbs->FindControls(vtMissing, vtMissing, variant_t(tag), vtMissing, &controls)) && controls != NULL)
		{
			int count = 0;
			controls->get_Count(&count);

			for (long i = count; i > 0; --i)
			{
				Office::CommandBarControlPtr control;
				if (SUCCEEDED(controls->get_Item(variant_t(i), &control)))
				{
					Office::_CommandBarButtonPtr button = control;

					if (button)
						UpdateItem(button);
				}
			}
		}
	}
}

void AddinUi::DestroyCommandBarsUI()
{
	m_buttons.RemoveAll();
}

HRESULT GetRibbonText(BSTR * RibbonXml)
{
	LPBYTE content = NULL;
	DWORD content_length = 0;

	ASSERT_RETURN_VALUE(LoadResourceFromModule(
		AfxGetInstanceHandle(), MAKEINTRESOURCE(IDR_RESOURCE_H), L"TEXT", &content, &content_length), S_FALSE);

	CMemFile mf(content, content_length);
	CTextFileRead rdr(&mf);

	CMapStringToString replacements;

	CString line;
	while (rdr.ReadLine(line))
	{
		CSimpleArray<CString> tokens;

		CString token;
		token.Preallocate(30);

		for (LPCWSTR pos = line; *pos; ++pos)
		{
			if (*pos == ' ' || *pos == '\t')
			{
				if (!token.IsEmpty())
				{
					tokens.Add(token);
					token = CString();
				}
			}
			else
			{
				token += *pos;
			}
		}

		if (!token.IsEmpty())
			tokens.Add(token);

		if (tokens.GetSize() != 3)
			continue;

		if (tokens[0] != "#define")
			continue;

		replacements[tokens[1]] = tokens[2];
	}

	CString ribbon = LoadTextFromModule(AfxGetInstanceHandle(), IDR_RIBBON);

	for (int pos = 0; pos < ribbon.GetLength(); ++pos)
	{
		if (ribbon[pos] != '{')
			continue;

		int endpos = ribbon.Find('}', pos);
		ASSERT_CONTINUE(endpos != -1);

		CString token = ribbon.Mid(pos+1, endpos-pos-1);
		CString token_found;

		ASSERT_CONTINUE(replacements.Lookup(token, token_found));	

		ribbon.Delete(pos, endpos-pos+1);
		ribbon.Insert(pos, token_found);

		pos += (token_found.GetLength()-1);
	}

	*RibbonXml = ribbon.AllocSysString();
	return S_OK;
}
