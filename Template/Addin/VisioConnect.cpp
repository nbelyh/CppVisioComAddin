// Connect.cpp : Implementation of CConnect

#include "stdafx.h"
#include "VisioConnect.h"
#include "AddinCommands.h"

#include <GdiPlus.h>
#pragma comment(lib,"gdiplus.lib")

/*------------------------------------------------------------------------------
    The interesting part - function to convert a image from resources into 
	IPictureDisp which is required by Visio, and optionally the mask
------------------------------------------------------------------------------*/

namespace {

	HRESULT CustomUiGetPng(LPCWSTR resourceId, bool large, IPictureDisp ** ppResultImage, IPictureDisp** ppResultMask)
	{
		HRESULT hr = S_OK;

		int sz = large ? 32 : 16;
		HICON hIcon = (HICON)LoadImage(_Module.GetResourceInstance(), resourceId, IMAGE_ICON, sz, sz, LR_SHARED);

		if (!hIcon)
			return HRESULT_FROM_WIN32(GetLastError());

		PICTDESC pd = { 0 };
		pd.cbSizeofstruct = sizeof(PICTDESC);
		pd.picType = PICTYPE_ICON;
		pd.icon.hicon = hIcon;

		if (ppResultMask == NULL) // no need for mask
			return OleCreatePictureIndirect(&pd, IID_IDispatch, TRUE, (LPVOID*)ppResultImage);

		using namespace Gdiplus;

		GdiplusStartupInput gdiplusStartupInput;
		ULONG_PTR gdiplusToken;

		gdiplusStartupInput.DebugEventCallback = NULL;
		gdiplusStartupInput.SuppressBackgroundThread = FALSE;
		gdiplusStartupInput.SuppressExternalCodecs = FALSE;
		gdiplusStartupInput.GdiplusVersion = 1;
		GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

		{
			Bitmap bitmap(hIcon);

			UINT w = bitmap.GetWidth();
			UINT h = bitmap.GetHeight();
			Rect r(0, 0, w, h);

			Bitmap bitmap_rgb(w, h, PixelFormat24bppRGB);
			Bitmap bitmap_msk(w, h, PixelFormat24bppRGB);

			BitmapData argb_bits;
			bitmap.LockBits(&r, ImageLockModeRead, PixelFormat32bppARGB, &argb_bits);

			BitmapData rgb_bits;
			bitmap_rgb.LockBits(&r, ImageLockModeWrite, PixelFormat24bppRGB, &rgb_bits);

			BitmapData msk_bits;
			bitmap_msk.LockBits(&r, ImageLockModeWrite, PixelFormat24bppRGB, &msk_bits);

			for (UINT y = 0; y < h; ++y)
			{
				for (UINT x = 0; x < w; ++x)
				{
					BYTE* idx_argb =
						static_cast<BYTE*>(argb_bits.Scan0) + (x + y*w) * 4;

					BYTE* idx_rgb =
						static_cast<BYTE*>(static_cast<BYTE*>(rgb_bits.Scan0) + (x + y*w) * 3);

					BYTE* idx_msk =
						static_cast<BYTE*>(static_cast<BYTE*>(msk_bits.Scan0) + (x + y*w) * 3);


					idx_rgb[0] = idx_argb[0];
					idx_rgb[1] = idx_argb[1];
					idx_rgb[2] = idx_argb[2];

					byte t = (idx_argb[3] < 128) ? 0xFF : 0x00;

					idx_msk[0] = t;
					idx_msk[1] = t;
					idx_msk[2] = t;
				}
			}

			bitmap.UnlockBits(&argb_bits);

			bitmap_rgb.UnlockBits(&rgb_bits);
			bitmap_msk.UnlockBits(&msk_bits);

			pd.picType = PICTYPE_BITMAP;

			bitmap_rgb.GetHBITMAP(0, &pd.bmp.hbitmap);
			hr = OleCreatePictureIndirect(&pd, IID_IDispatch, TRUE, (LPVOID*)ppResultImage);

			bitmap_msk.GetHBITMAP(0, &pd.bmp.hbitmap);
			hr = OleCreatePictureIndirect(&pd, IID_IDispatch, TRUE, (LPVOID*)ppResultMask);
		}

		GdiplusShutdown(gdiplusToken);
		return hr;
	}

	_ATL_FUNC_INFO ClickEventInfo = { CC_STDCALL, VT_EMPTY, 2, { VT_DISPATCH, VT_BOOL | VT_BYREF } };

	// event sink to handle button click events from Visio.
	class ClickEventSink :
		public IDispEventSimpleImpl<1, ClickEventSink, &__uuidof(Office::_CommandBarButtonEvents)>
	{
	public:
		ClickEventSink(CAddinCommands* pCommands, IUnknownPtr pUnk) 
			: m_pCommands(pCommands)
			, m_pUnk(pUnk)
		{
			DispEventAdvise(m_pUnk);
		}

		~ClickEventSink()
		{
			DispEventUnadvise(m_pUnk);
		}

		CAddinCommands* m_pCommands;

		// keep the reference to the item itself, otherwise VISIO may destroys it
		IUnknownPtr m_pUnk;

		// the event handler itself. 
		void __stdcall  OnClick(IDispatch* pButton, VARIANT_BOOL* /* pCancel */)
		{
			Office::CommandBarControlPtr control = pButton;
			m_pCommands->OnCommand(control->Tag);
		}

		BEGIN_SINK_MAP(ClickEventSink)
			SINK_ENTRY_INFO(1, __uuidof(Office::_CommandBarButtonEvents), 1, &ClickEventSink::OnClick, &ClickEventInfo)
		END_SINK_MAP()
	};
}

struct CVisioConnect::Impl
{
	/*------------------------------------------------------------------------------
		Adds a single button using the CommandBars interface
	------------------------------------------------------------------------------*/

	void InstallButton(Office::CommandBarPtr commandBar, LPCWSTR tag, LPCWSTR caption)
	{
		Office::_CommandBarButtonPtr button = commandBar->FindControl(vtMissing, vtMissing, tag);

		if (!button)
		{
			button = commandBar->Controls->Add(int(Office::msoControlButton));
			button->Tag = tag;
			button->Caption = caption;
		}

		UINT iconId = m_pCommands->GetCommandIconId(tag);
		
		if (iconId >= 0)
		{
			IPictureDispPtr disp_picture;
			IPictureDispPtr disp_mask;

			CustomUiGetPng(MAKEINTRESOURCE(iconId), false, &disp_picture, &disp_mask);

			button->put_Picture(disp_picture);
			button->put_Mask(disp_mask);
		}

		ClickEventSink* sink = new ClickEventSink(m_pCommands, button);
		m_buttons.Add(tag, sink);
	}

	/*------------------------------------------------------------------------------
		Constructs a toolbar using the CommandBars model
	------------------------------------------------------------------------------*/

	void InstallToolbar()
	{
		struct Builder : IToolbarBuilder
		{
			Impl* m_this;
			Office::CommandBarPtr m_cb;

			virtual void AddToolbar(LPCWSTR name)
			{
				Office::_CommandBarsPtr cbs = m_this->m_visio->CommandBars;

				variant_t vName = name;
				if (FAILED(cbs->get_Item(vName, &m_cb)))
					m_cb = cbs->Add(vName);

				m_cb->Visible = VARIANT_TRUE;
			}

			virtual void AddButton(LPCWSTR cmd)
			{
				bstr_t label = m_this->m_pCommands->GetCommandLabel(cmd);
				m_this->InstallButton(m_cb, cmd, label);
			}

		} builder;
		
		builder.m_this = this;

		m_pCommands->BuildToolbar(&builder);
	}

	/*------------------------------------------------------------------------------
		On connection, built a toolbar (for old Visio; new Visio uses ribbon)
	------------------------------------------------------------------------------*/
	void Connect(Visio::IVApplicationPtr visio)
	{
		m_visio = visio;

		m_pCommands = new CAddinCommands(m_visio);

		LPCWSTR version = visio->Version;

		if (StrToInt(version) < 14)
			InstallToolbar();
	}

	/*------------------------------------------------------------------------------
		On disconnection, remove all our buttons and event handlers
	------------------------------------------------------------------------------*/
	void Disconnect()
	{
		for (int i = 0; i < m_buttons.GetSize(); ++i)
			delete m_buttons.GetValueAt(i);

		m_buttons.RemoveAll();

		delete m_pCommands;

		m_visio = NULL;
	}

	CAddinCommands* m_pCommands;
	Visio::IVApplicationPtr m_visio;
	CSimpleMap<bstr_t, ClickEventSink*> m_buttons;
};

/*------------------------------------------------------------------------------
    If visio 2010 or older (ribbon), do not create toolbars, otherwise do
------------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnConnection(IDispatch *pApplication, ext_ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	UNREFERENCED_PARAMETER(pAddInInst);

	m_impl->Connect(pApplication);

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
	m_impl->Disconnect();

	return S_OK;
}

STDMETHODIMP CVisioConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

/*------------------------------------------------------------------------------
    Used by Visio 2010 to query addin for custom UI
------------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::raw_GetCustomUI(BSTR ribbonID, BSTR * pRibbonXml)
{
	UNREFERENCED_PARAMETER(ribbonID);

	HMODULE hResources = _Module.GetResourceInstance();

	HRSRC rc = ::FindResource(
		hResources, MAKEINTRESOURCE(IDR_RIBBON), L"TEXTFILE");

	LPWSTR content = static_cast<LPWSTR>(
		::LockResource(::LoadResource(hResources, rc)));

	DWORD content_length = 
		::SizeofResource(hResources, rc);

	*pRibbonXml = SysAllocStringLen(content, content_length / 2);

	return S_OK;
}

/*------------------------------------------------------------------------------
    Used by Visio ribbon to query for an image
------------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnRibbonGetImage(IDispatch *pControl, IPictureDisp ** ppdispImage)
{
	Office::IRibbonControlPtr ctrl = pControl;
	UINT imageId = m_impl->m_pCommands->GetCommandIconId(ctrl->Tag);

	if (imageId >= 0)
		return CustomUiGetPng(MAKEINTRESOURCE(imageId), true, ppdispImage, NULL);

	return S_FALSE;
}

/*------------------------------------------------------------------------------
    Called by Visio 2010 ribbon interface
------------------------------------------------------------------------------*/

STDMETHODIMP CVisioConnect::OnRibbonGetLabel(IDispatch *pControl, BSTR* pLabel)
{
	Office::IRibbonControlPtr ctrl = pControl;
	bstr_t label = m_impl->m_pCommands->GetCommandLabel(ctrl->Tag);
	*pLabel = label.Detach();
	return S_OK;
}

STDMETHODIMP CVisioConnect::OnRibbonButtonClicked (IDispatch * pControl)
{
	Office::IRibbonControlPtr ctrl = pControl;
	m_impl->m_pCommands->OnCommand(ctrl->Tag);
	return S_OK;
}

/*------------------------------------------------------------------------------
    The worker function; just show a message box
------------------------------------------------------------------------------*/

CVisioConnect::CVisioConnect()
	: m_impl(new Impl())
{
}

CVisioConnect::~CVisioConnect()
{
	delete m_impl;
}

BEGIN_OBJECT_MAP(ObjectMap)
	OBJECT_ENTRY(__uuidof(VisioConnect), CVisioConnect)
END_OBJECT_MAP()
