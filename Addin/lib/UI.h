
#pragma once

#include <atlcoll.h>

extern _ATL_FUNC_INFO ClickEventInfo;

// event sink to handle button click events from Visio.
// Finally this simple "click" event is used as the least evil method.
class ClickEventRedirector :
	public IDispEventSimpleImpl<1, ClickEventRedirector, &__uuidof(Office::_CommandBarButtonEvents)>
{
public:
	ClickEventRedirector(IUnknownPtr punk);
	~ClickEventRedirector();

	// the event handler itself. 
	// Just redirect to the global command processor ("parameter" of the button contains the command itself)
	void __stdcall OnClick(IDispatch* pButton, VARIANT_BOOL* pCancel);

	// keep the reference to the item itself, otherwise VISIO destroys it for some unknown reason
	// probably because of it's "double-nature" user interface, and events are never fired.
	IUnknownPtr m_punk;

	BEGIN_SINK_MAP(ClickEventRedirector)
		SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents), 1, &ClickEventRedirector::OnClick, &ClickEventInfo)
	END_SINK_MAP()
};

struct AddinUi
{
	void CreateCommandBarsUI(IVApplicationPtr app);
	void DestroyCommandBarsUI();

private:
	void InitializeItem( Office::CommandBarControlPtr item, UINT command_id );
	void FillMenuItems( long position, Office::CommandBarControlsPtr menu_items, CMenu* popup_menu );
	void FillMenu( long position, Office::CommandBarControlsPtr cbs, UINT menu_id );
	Office::CommandBarPopupPtr CreateFrameMenu(Office::CommandBarControlsPtr menus);
private:
	CAtlArray<ClickEventRedirector*> m_buttons;
};

HRESULT GetRibbonText(BSTR * RibbonXml);
