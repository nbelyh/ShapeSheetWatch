// Connect.h : Declaration of the CVisioConnect

#pragma once

using namespace Office;
using namespace AddInDesignerObjects;

// CVisioConnect
class ATL_NO_VTABLE CVisioConnect : 
	public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CVisioConnect, &__uuidof(VisioConnect)>
	, public IDispatchImpl<ICallbackInterface, &__uuidof(ICallbackInterface), &LIBID_AddinLib, 1, 0>
	, public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
	, public IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), 12, 0>
{
public:
	CVisioConnect();
	~CVisioConnect();

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_NOT_AGGREGATABLE(CVisioConnect)

BEGIN_COM_MAP(CVisioConnect)
	COM_INTERFACE_ENTRY2(IDispatch, ICallbackInterface)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
	COM_INTERFACE_ENTRY(IRibbonExtensibility)
	COM_INTERFACE_ENTRY(ICallbackInterface)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();
	void FinalRelease();

public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	//IRibbonExtensibility implementation:
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);

	STDMETHOD(OnRibbonButtonClicked)(IDispatch * RibbonControl);

	STDMETHOD(OnRibbonLoad)(IDispatch* disp);
	STDMETHOD(OnRibbonLoadImage)(BSTR pbstrImageId, IPictureDisp ** ppdispImage);

	STDMETHOD(IsRibbonButtonEnabled)(IDispatch * RibbonControl, VARIANT_BOOL* pResult);
	STDMETHOD(IsRibbonButtonVisible)(IDispatch * RibbonControl, VARIANT_BOOL* pResult);
	STDMETHOD(OnRibbonCheckboxClicked)(IDispatch * RibbonControl, VARIANT_BOOL* pResult);
	STDMETHOD(IsRibbonButtonPressed)(IDispatch * RibbonControl, VARIANT_BOOL* pResult);

	STDMETHOD(GetRibbonLabel)(IDispatch *pControl, BSTR *pbstrLabel);
	STDMETHOD(GetRibbonImage)(IDispatch *pControl, IPictureDisp ** ppdispImage);
private:
	struct Impl;
	Impl* m_impl;

};

OBJECT_ENTRY_AUTO(__uuidof(VisioConnect), CVisioConnect)
