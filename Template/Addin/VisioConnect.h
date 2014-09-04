// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols

#include "AddIn_i.h"

// CConnect
class ATL_NO_VTABLE CVisioConnect :
	public CComObjectRootEx<CComSingleThreadModel>
	, public CComCoClass<CVisioConnect, &CLSID_VisioConnect>
	, public IDispatchImpl<ICallbackInterface, &__uuidof(ICallbackInterface), &LIBID_AddinLib, 1, 0>
	, public IDispatchImpl<IDTExtensibility2, &__uuidof(IDTExtensibility2)>
	, public IDispatchImpl<Office::IRibbonExtensibility, &__uuidof(Office::IRibbonExtensibility)>
{
public:
	CVisioConnect();
	virtual ~CVisioConnect();

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_NOT_AGGREGATABLE(CVisioConnect)

BEGIN_COM_MAP(CVisioConnect)
	COM_INTERFACE_ENTRY2(IDispatch, ICallbackInterface)
	COM_INTERFACE_ENTRY(IDTExtensibility2)
	COM_INTERFACE_ENTRY(Office::IRibbonExtensibility)
	COM_INTERFACE_ENTRY(ICallbackInterface)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * application, ext_ConnectMode connectMode, IDispatch *pAddInInst, SAFEARRAY **custom);

	STDMETHOD(OnDisconnection)(ext_DisconnectMode removeMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	//IRibbonExtensibility implementation:
	STDMETHOD(raw_GetCustomUI)(BSTR RibbonID, BSTR * RibbonString);

	STDMETHOD(OnRibbonButtonClicked)(IDispatch * pControl);
	STDMETHOD(OnRibbonGetLabel)(IDispatch *pControl, BSTR* pLabel);
	STDMETHOD(OnRibbonGetImage)(IDispatch *pControl, IPictureDisp ** ppdispImage);

	void OnButton(LPCWSTR tag);

private:

	struct Impl;
	Impl* m_impl;
};

OBJECT_ENTRY_AUTO(__uuidof(VisioConnect), CVisioConnect)
