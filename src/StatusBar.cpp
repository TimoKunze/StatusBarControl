// StatusBar.cpp: Superclasses msctls_statusbar32.

#include "stdafx.h"
#include "StatusBar.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC StatusBar::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


StatusBar::StatusBar()
{
	SIZEL size = {100, 23};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	properties.font.InitializePropertyWatcher(this, DISPID_STATBAR_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_STATBAR_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	panelUnderMouse = -2;
	pToolTipBuffer = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (MAX_PATH + 1) * sizeof(WCHAR));

	properties.panels = NULL;
	CComObject<StatusBarPanels>::CreateInstance(&properties.panels);
	properties.panels->AddRef();

	properties.simplePanel = NULL;
	CComObject<StatusBarPanel>::CreateInstance(&properties.simplePanel);
	properties.simplePanel->AddRef();
	properties.simplePanel->properties.panelIndex = SB_SIMPLEID;

	ATLASSUME(properties.panels);
	ATLASSUME(properties.simplePanel);

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}
}

StatusBar::~StatusBar()
{
	if(pToolTipBuffer) {
		HeapFree(GetProcessHeap(), 0, pToolTipBuffer);
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP StatusBar::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IStatusBar, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP StatusBar::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	ATLASSUME(pPropertyBag);
	if(!pPropertyBag) {
		return E_POINTER;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	if(SUCCEEDED(pPropertyBag->Read(OLESTR("Version"), &propertyValue, pErrorLog))) {
		if(propertyValue.lVal > 0x0102) {
			return E_FAIL;
		}
	}
	LONG version = propertyValue.lVal;

	propertyValue.vt = VT_UI4;
	SIZE extent = {0};
	HRESULT hr = pPropertyBag->Read(OLESTR("_cx"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	extent.cx = propertyValue.ulVal;
	hr = pPropertyBag->Read(OLESTR("_cy"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	extent.cy = propertyValue.ulVal;

	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("Appearance"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	hr = pPropertyBag->Read(OLESTR("BackColor"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	hr = pPropertyBag->Read(OLESTR("BorderStyle"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	CComBSTR customCapsLockText;
	CComBSTR customInsertKeyText;
	CComBSTR customKanaLockText;
	CComBSTR customNumLockText;
	CComBSTR customScrollLockText;
	if(version < 0x0102)
	{
		propertyValue.vt = VT_BSTR;
		hr = pPropertyBag->Read(OLESTR("CustomCapsLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		customCapsLockText = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
		hr = pPropertyBag->Read(OLESTR("CustomInsertKeyText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		customInsertKeyText = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
		hr = pPropertyBag->Read(OLESTR("CustomKanaLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		customKanaLockText = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
		hr = pPropertyBag->Read(OLESTR("CustomNumLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		customNumLockText = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
		hr = pPropertyBag->Read(OLESTR("CustomScrollLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		customScrollLockText = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
	} else {
		propertyValue.vt = VT_ARRAY | VT_UI1;
		hr = pPropertyBag->Read(OLESTR("CustomCapsLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			customCapsLockText.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
		hr = pPropertyBag->Read(OLESTR("CustomInsertKeyText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			customInsertKeyText.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
		hr = pPropertyBag->Read(OLESTR("CustomKanaLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			customKanaLockText.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
		hr = pPropertyBag->Read(OLESTR("CustomNumLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			customNumLockText.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
		hr = pPropertyBag->Read(OLESTR("CustomScrollLockText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			customScrollLockText.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
	}
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("DisabledEvents"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("DontRedraw"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	hr = pPropertyBag->Read(OLESTR("Enabled"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	hr = pPropertyBag->Read(OLESTR("Font"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont;
	if(propertyValue.pdispVal) {
		propertyValue.pdispVal->QueryInterface(IID_PPV_ARGS(&pFont));
	}
	VariantClear(&propertyValue);

	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("ForceSizeGripperDisplay"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL forceSizeGripperDisplay = propertyValue.boolVal;
	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("HoverTime"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	hr = pPropertyBag->Read(OLESTR("MinimumHeight"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	OLE_YSIZE_PIXELS minimumHeight = propertyValue.lVal;

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	hr = pPropertyBag->Read(OLESTR("MouseIcon"), &propertyValue, pErrorLog);
	CComPtr<IPictureDisp> pMouseIcon;
	if(propertyValue.pdispVal) {
		propertyValue.pdispVal->QueryInterface(IID_PPV_ARGS(&pMouseIcon));
	}
	VariantClear(&propertyValue);

	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("MousePointer"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	CComObject<StatusBarPanels>* pPanels = NULL;
	CComObject<StatusBarPanels>::CreateInstance(&pPanels);
	pPanels->AddRef();
	pPanels->SetOwner(this);
	pPanels->QueryInterface(IID_PPV_ARGS(&propertyValue.pdispVal));
	hr = pPropertyBag->Read(OLESTR("Panels"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VariantClear(&propertyValue);

	propertyValue.vt = VT_I4;
	hr = pPropertyBag->Read(OLESTR("PanelToAutoSize"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	CComPtr<IStatusBarPanel> panelToAutoSize = NULL;
	if(propertyValue.lVal != -1) {
		pPanels->get_Item(propertyValue.lVal, &panelToAutoSize);
	}
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("RegisterForOLEDragDrop"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	hr = pPropertyBag->Read(OLESTR("RightToLeftLayout"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL rightToLeftLayout = propertyValue.boolVal;
	hr = pPropertyBag->Read(OLESTR("ShowToolTips"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL showToolTips = propertyValue.boolVal;
	hr = pPropertyBag->Read(OLESTR("SimpleMode"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL simpleMode = propertyValue.boolVal;

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	CComObject<StatusBarPanel>* pSimplePanel = NULL;
	CComObject<StatusBarPanel>::CreateInstance(&pSimplePanel);
	pSimplePanel->AddRef();
	pSimplePanel->SetOwner(this);
	pSimplePanel->properties.panelIndex = SB_SIMPLEID;
	pSimplePanel->QueryInterface(IID_PPV_ARGS(&propertyValue.pdispVal));
	hr = pPropertyBag->Read(OLESTR("SimplePanel"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VariantClear(&propertyValue);

	hr = pPropertyBag->Read(OLESTR("SupportOLEDragImages"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	hr = pPropertyBag->Read(OLESTR("UseSystemFont"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;

	SetExtent(DVASPECT_CONTENT, &extent);
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomCapsLockText(customCapsLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomInsertKeyText(customInsertKeyText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomKanaLockText(customKanaLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomNumLockText(customNumLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomScrollLockText(customScrollLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForceSizeGripperDisplay(forceSizeGripperDisplay);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinimumHeight(minimumHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	CComObject<StatusBarPanels>* pOldPanels = properties.panels;
	properties.panels = pPanels;
	hr = putref_PanelToAutoSize(panelToAutoSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeftLayout(rightToLeftLayout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowToolTips(showToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SimpleMode(simpleMode);
	if(FAILED(hr)) {
		return hr;
	}
	CComObject<StatusBarPanel>* pOldSimplePanel = properties.simplePanel;
	properties.simplePanel = pSimplePanel;
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	if(IsWindow()) {
		if(properties.simpleMode) {
			EnterInternalSBSIMPLESection();
			SendMessage(SB_SIMPLE, FALSE, 0);
		}
		ApplyPanelWidths(TRUE);
		if(properties.simpleMode) {
			SendMessage(SB_SIMPLE, TRUE, 0);
			LeaveInternalSBSIMPLESection();
		}
		properties.panels->ApplySettings(FALSE, TRUE);
	}

	if(pOldPanels) {
		pOldPanels->Release();
	}
	if(pOldSimplePanel) {
		pOldSimplePanel->Release();
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP StatusBar::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/)
{
	ATLASSUME(pPropertyBag);
	if(!pPropertyBag) {
		return E_POINTER;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	HRESULT hr = S_OK;
	propertyValue.vt = VT_I4;
	propertyValue.lVal = 0x0102;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Version"), &propertyValue))) {
		return hr;
	}

	propertyValue.vt = VT_UI4;
	propertyValue.ulVal = m_sizeExtent.cx;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("_cx"), &propertyValue))) {
		return hr;
	}
	propertyValue.ulVal = m_sizeExtent.cy;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("_cy"), &propertyValue))) {
		return hr;
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Appearance"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("BackColor"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("BorderStyle"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.customCapsLockText.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("CustomCapsLockText"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.customInsertKeyText.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("CustomInsertKeyText"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.customKanaLockText.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("CustomKanaLockText"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.customNumLockText.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("CustomNumLockText"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.customScrollLockText.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("CustomScrollLockText"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DisabledEvents"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("DontRedraw"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Enabled"), &propertyValue))) {
		return hr;
	}

	if(properties.font.pFontDisp) {
		propertyValue.vt = VT_DISPATCH;
		propertyValue.pdispVal = NULL;
		properties.font.pFontDisp->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));
		if(FAILED(hr = pPropertyBag->Write(OLESTR("Font"), &propertyValue))) {
			VariantClear(&propertyValue);
			return hr;
		}
		VariantClear(&propertyValue);
	}

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.forceSizeGripperDisplay);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ForceSizeGripperDisplay"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("HoverTime"), &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minimumHeight;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("MinimumHeight"), &propertyValue))) {
		return hr;
	}

	if(properties.mouseIcon.pPictureDisp) {
		propertyValue.vt = VT_DISPATCH;
		propertyValue.pdispVal = NULL;
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));
		if(FAILED(hr = pPropertyBag->Write(OLESTR("MouseIcon"), &propertyValue))) {
			VariantClear(&propertyValue);
			return hr;
		}
		VariantClear(&propertyValue);
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("MousePointer"), &propertyValue))) {
		return hr;
	}

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	properties.panels->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Panels"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);

	propertyValue.vt = VT_I4;
	if(properties.panelToAutoSize) {
		properties.panelToAutoSize->get_Index(&propertyValue.lVal);
	} else {
		propertyValue.lVal = -1;
	}
	if(FAILED(hr = pPropertyBag->Write(OLESTR("PanelToAutoSize"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("RegisterForOLEDragDrop"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("RightToLeftLayout"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showToolTips);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ShowToolTips"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.simpleMode);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("SimpleMode"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("SupportOLEDragImages"), &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("UseSystemFont"), &propertyValue))) {
		return hr;
	}

	propertyValue.vt = VT_DISPATCH;
	propertyValue.pdispVal = NULL;
	properties.simplePanel->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));
	if(FAILED(hr = pPropertyBag->Write(OLESTR("SimplePanel"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 8 VT_I4 properties...
	pSize->QuadPart += 8 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 9 VT_BOOL properties...
	pSize->QuadPart += 9 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	// ...and 5 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.customCapsLockText.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.customInsertKeyText.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.customKanaLockText.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.customNumLockText.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.customScrollLockText.ByteLength() + sizeof(OLECHAR);

	// ...and 4 VT_DISPATCH properties
	pSize->QuadPart += 4 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(FAILED(properties.panels->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
		properties.panels->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(FAILED(properties.simplePanel->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
		properties.simplePanel->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP StatusBar::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x0D0D0D0D/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0101) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR customCapsLockText;
	if(FAILED(hr = customCapsLockText.ReadFromStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR customInsertKeyText;
	if(FAILED(hr = customInsertKeyText.ReadFromStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR customKanaLockText;
	if(FAILED(hr = customKanaLockText.ReadFromStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR customNumLockText;
	if(FAILED(hr = customNumLockText.ReadFromStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR customScrollLockText;
	if(FAILED(hr = customScrollLockText.ReadFromStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL forceSizeGripperDisplay = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG minimumHeight = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CLSID clsid = CLSID_NULL;
	if(FAILED(hr = ReadClassStm(pStream, &clsid))) {
		return hr;
	}
	CComObject<StatusBarPanels>* pPanels = NULL;
	CComObject<StatusBarPanels>::CreateInstance(&pPanels);
	pPanels->AddRef();
	pPanels->SetOwner(this);
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(FAILED(hr = pPanels->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
		if(FAILED(hr = pPanels->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			return hr;
		}
	}
	if(pStreamInit) {
		/*CLSID storedCLSID;
		if(FAILED(hr = pStreamInit->GetClassID(&storedCLSID))) {
			return hr;
		}*/
		// {CCA75315-B100-4b5f-80F6-8DFE616F8FDB}
		CLSID CLSID_StatusBarPanelsW = {0xCCA75315, 0xB100, 0x4b5f, {0x80, 0xF6, 0x8D, 0xFE, 0x61, 0x6F, 0x8F, 0xDB}};
		// {9A38B78A-CE98-440b-B1F6-1CE994432F2B}
		CLSID CLSID_StatusBarPanelsA = {0x9A38B78A, 0xCE98, 0x440b, {0xB1, 0xF6, 0x1C, 0xE9, 0x94, 0x43, 0x2F, 0x2B}};
		if((clsid == CLSID_StatusBarPanelsA) || (clsid == CLSID_StatusBarPanelsW)) {
			if(FAILED(hr = pStreamInit->Load(pStream))) {
				return hr;
			}
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	CComPtr<IStatusBarPanel> panelToAutoSize = NULL;
	if(propertyValue.lVal != -1) {
		pPanels->get_Item(propertyValue.lVal, &panelToAutoSize);
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL registerForOLEDragDrop = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rightToLeftLayout = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL showToolTips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL simpleMode = propertyValue.boolVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	if(FAILED(hr = ReadClassStm(pStream, &clsid))) {
		return hr;
	}
	CComObject<StatusBarPanel>* pSimplePanel = NULL;
	CComObject<StatusBarPanel>::CreateInstance(&pSimplePanel);
	pSimplePanel->AddRef();
	pSimplePanel->SetOwner(this);
	pSimplePanel->properties.panelIndex = SB_SIMPLEID;
	pStreamInit = NULL;
	if(FAILED(hr = pSimplePanel->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
		if(FAILED(hr = pSimplePanel->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			return hr;
		}
	}
	if(pStreamInit) {
		/*CLSID storedCLSID;
		if(FAILED(hr = pStreamInit->GetClassID(&storedCLSID))) {
			return hr;
		}*/
		// {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B}
		CLSID CLSID_StatusBarPanelW = {0xCB0F173F, 0x9E1F, 0x4365, {0xBF, 0x3C, 0x6C, 0xC5, 0x2F, 0x8C, 0x26, 0x8B}};
		// {72208630-41D0-479f-B075-7727847698DE}
		CLSID CLSID_StatusBarPanelA = {0x72208630, 0x41D0, 0x479f, {0xB0, 0x75, 0x77, 0x27, 0x84, 0x76, 0x98, 0xDE}};
		if((clsid == CLSID_StatusBarPanelA) || (clsid == CLSID_StatusBarPanelW)) {
			if(FAILED(hr = pStreamInit->Load(pStream))) {
				return hr;
			}
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;


	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomCapsLockText(customCapsLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomInsertKeyText(customInsertKeyText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomKanaLockText(customKanaLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomNumLockText(customNumLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_CustomScrollLockText(customScrollLockText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForceSizeGripperDisplay(forceSizeGripperDisplay);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinimumHeight(minimumHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	CComObject<StatusBarPanels>* pOldPanels = properties.panels;
	properties.panels = pPanels;
	hr = putref_PanelToAutoSize(panelToAutoSize);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RightToLeftLayout(rightToLeftLayout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShowToolTips(showToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SimpleMode(simpleMode);
	if(FAILED(hr)) {
		return hr;
	}
	CComObject<StatusBarPanel>* pOldSimplePanel = properties.simplePanel;
	properties.simplePanel = pSimplePanel;
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}

	if(IsWindow()) {
		if(properties.simpleMode) {
			EnterInternalSBSIMPLESection();
			SendMessage(SB_SIMPLE, FALSE, 0);
		}
		ApplyPanelWidths(TRUE);
		if(properties.simpleMode) {
			SendMessage(SB_SIMPLE, TRUE, 0);
			LeaveInternalSBSIMPLESection();
		}
		properties.panels->ApplySettings(FALSE, TRUE);
	}

	if(pOldPanels) {
		pOldPanels->Release();
	}
	if(pOldSimplePanel) {
		pOldSimplePanel->Release();
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP StatusBar::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x0D0D0D0D/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0101;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	VARTYPE vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.customCapsLockText.WriteToStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.customInsertKeyText.WriteToStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.customKanaLockText.WriteToStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.customNumLockText.WriteToStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.customScrollLockText.WriteToStream(pStream))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.forceSizeGripperDisplay);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.minimumHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStreamInit> pPersistInit = NULL;
	if(FAILED(hr = properties.panels->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistInit)))) {
		if(FAILED(hr = properties.panels->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pPersistInit)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistInit) {
		CLSID clsid = CLSID_NULL;
		if(FAILED(hr = pPersistInit->GetClassID(&clsid))) {
			return hr;
		}
		if(FAILED(hr = WriteClassStm(pStream, clsid))) {
			return hr;
		}
		if(FAILED(hr = pPersistInit->Save(pStream, clearDirtyFlag))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	if(properties.panelToAutoSize) {
		properties.panelToAutoSize->get_Index(&propertyValue.lVal);
	} else {
		propertyValue.lVal = -1;
	}
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showToolTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.simpleMode);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistInit = NULL;
	if(FAILED(hr = properties.simplePanel->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistInit)))) {
		if(FAILED(hr = properties.simplePanel->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pPersistInit)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistInit) {
		CLSID clsid = CLSID_NULL;
		if(FAILED(hr = pPersistInit->GetClassID(&clsid))) {
			return hr;
		}
		if(FAILED(hr = WriteClassStm(pStream, clsid))) {
			return hr;
		}
		if(FAILED(hr = pPersistInit->Save(pStream, clearDirtyFlag))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


StatusBar::Properties::Properties()
{
	panels = NULL;
	simplePanel = NULL;
	ResetToDefaults();
}

StatusBar::Properties::~Properties()
{
	if(panels) {
		panels->Release();
	}
	if(simplePanel) {
		simplePanel->Release();
	}
}


HWND StatusBar::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_BAR_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<StatusBar>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT StatusBar::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("StatusBar ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void StatusBar::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	if(pToolTipBuffer) {
		HeapFree(GetProcessHeap(), 0, pToolTipBuffer);
		pToolTipBuffer = NULL;
	}
	Release();
}

STDMETHODIMP StatusBar::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	properties.panels->SetOwner(this);
	properties.simplePanel->SetOwner(this);

	HRESULT hr = CComControl<StatusBar>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP StatusBar::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP StatusBar::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP StatusBar::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP StatusBar::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_STATBAR_APPEARANCE:
		case DISPID_STATBAR_BORDERSTYLE:
		case DISPID_STATBAR_CUSTOMCAPSLOCKTEXT:
		case DISPID_STATBAR_CUSTOMINSERTKEYTEXT:
		case DISPID_STATBAR_CUSTOMKANALOCKTEXT:
		case DISPID_STATBAR_CUSTOMNUMLOCKTEXT:
		case DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT:
		case DISPID_STATBAR_MINIMUMHEIGHT:
		case DISPID_STATBAR_MOUSEICON:
		case DISPID_STATBAR_MOUSEPOINTER:
		case DISPID_STATBAR_RIGHTTOLEFTLAYOUT:
		case DISPID_STATBAR_SIMPLEMODE:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_STATBAR_DISABLEDEVENTS:
		case DISPID_STATBAR_DONTREDRAW:
		case DISPID_STATBAR_FORCESIZEGRIPPERDISPLAY:
		case DISPID_STATBAR_HOVERTIME:
		case DISPID_STATBAR_PANELTOAUTOSIZE:
		case DISPID_STATBAR_SHOWTOOLTIPS:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_STATBAR_BACKCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_STATBAR_APPID:
		case DISPID_STATBAR_APPNAME:
		case DISPID_STATBAR_APPSHORTNAME:
		case DISPID_STATBAR_BUILD:
		case DISPID_STATBAR_CHARSET:
		case DISPID_STATBAR_HWND:
		case DISPID_STATBAR_ISRELEASE:
		case DISPID_STATBAR_PROGRAMMER:
		case DISPID_STATBAR_TESTER:
		case DISPID_STATBAR_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_STATBAR_REGISTERFOROLEDRAGDROP:
		case DISPID_STATBAR_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_STATBAR_FONT:
		case DISPID_STATBAR_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_STATBAR_PANELS:
		case DISPID_STATBAR_SIMPLEPANEL:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_STATBAR_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString StatusBar::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString StatusBar::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString StatusBar::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=StatusBar&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString StatusBar::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString StatusBar::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString StatusBar::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP StatusBar::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_STATBAR_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_STATBAR_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_STATBAR_CUSTOMCAPSLOCKTEXT:
		case DISPID_STATBAR_CUSTOMINSERTKEYTEXT:
		case DISPID_STATBAR_CUSTOMKANALOCKTEXT:
		case DISPID_STATBAR_CUSTOMNUMLOCKTEXT:
		case DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT:
			#ifdef UNICODE
				description = TEXT("(Unicode Text)");
			#else
				description = TEXT("(ANSI Text)");
			#endif
			hr = S_OK;
			break;
		case DISPID_STATBAR_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<StatusBar>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP StatusBar::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_STATBAR_BORDERSTYLE:
			c = 2;
			break;
		case DISPID_STATBAR_APPEARANCE:
			c = 3;
			break;
		case DISPID_STATBAR_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<StatusBar>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_STATBAR_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBar::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_STATBAR_APPEARANCE:
		case DISPID_STATBAR_BORDERSTYLE:
		case DISPID_STATBAR_MOUSEPOINTER:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<StatusBar>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP StatusBar::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	if(!pPropertyPage) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_STATBAR_CUSTOMCAPSLOCKTEXT:
		case DISPID_STATBAR_CUSTOMINSERTKEYTEXT:
		case DISPID_STATBAR_CUSTOMKANALOCKTEXT:
		case DISPID_STATBAR_CUSTOMNUMLOCKTEXT:
		case DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT:
			*pPropertyPage = CLSID_StringProperties;
			return S_OK;
		default:
			*pPropertyPage = CLSID_NULL;
			return PERPROP_E_NOPAGEAVAILABLE;
	}
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT StatusBar::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_STATBAR_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_STATBAR_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_STATBAR_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP StatusBar::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 6;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StringProperties;
		pPropertyPages->pElems[2] = CLSID_PanelProperties;
		pPropertyPages->pElems[3] = CLSID_StockColorPage;
		pPropertyPages->pElems[4] = CLSID_StockFontPage;
		pPropertyPages->pElems[5] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP StatusBar::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<StatusBar>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP StatusBar::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP StatusBar::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;
	pControlInfo->cAccel = 0;
	pControlInfo->dwFlags = 0;
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT StatusBar::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT StatusBar::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_STATBAR_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT StatusBar::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP StatusBar::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP StatusBar::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_APPEARANCE);

	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_STATBAR_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 13;
	return S_OK;
}

STDMETHODIMP StatusBar::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"StatusBar");
	return S_OK;
}

STDMETHODIMP StatusBar::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"StatBar");
	return S_OK;
}

STDMETHODIMP StatusBar::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		// there's no SB_GETBKCOLOR, but SB_SETBKCOLOR returns the previous color
		COLORREF color = SendMessage(SB_SETBKCOLOR, 0, 0);
		SendMessage(SB_SETBKCOLOR, 0, color);

		if(color == CLR_DEFAULT) {
			properties.backColor = CLR_NONE;
		} else if(color != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = color;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP StatusBar::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_BACKCOLOR);

	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(SB_SETBKCOLOR, 0, (properties.backColor == CLR_NONE ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.backColor)));
		}
		FireOnChanged(DISPID_STATBAR_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP StatusBar::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_BORDERSTYLE);

	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_STATBAR_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP StatusBar::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP StatusBar::get_CustomCapsLockText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.customCapsLockText.Copy();
	return S_OK;
}

STDMETHODIMP StatusBar::put_CustomCapsLockText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_CUSTOMCAPSLOCKTEXT);

	if(properties.customCapsLockText != newValue) {
		properties.customCapsLockText.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.panels->ApplySettings(FALSE, TRUE);
		}
		FireOnChanged(DISPID_STATBAR_CUSTOMCAPSLOCKTEXT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_CustomInsertKeyText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.customInsertKeyText.Copy();
	return S_OK;
}

STDMETHODIMP StatusBar::put_CustomInsertKeyText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_CUSTOMINSERTKEYTEXT);

	if(properties.customInsertKeyText != newValue) {
		properties.customInsertKeyText.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.panels->ApplySettings(FALSE, TRUE);
		}
		FireOnChanged(DISPID_STATBAR_CUSTOMINSERTKEYTEXT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_CustomKanaLockText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.customKanaLockText.Copy();
	return S_OK;
}

STDMETHODIMP StatusBar::put_CustomKanaLockText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_CUSTOMKANALOCKTEXT);

	if(properties.customKanaLockText != newValue) {
		properties.customKanaLockText.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.panels->ApplySettings(FALSE, TRUE);
		}
		FireOnChanged(DISPID_STATBAR_CUSTOMKANALOCKTEXT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_CustomNumLockText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.customNumLockText.Copy();
	return S_OK;
}

STDMETHODIMP StatusBar::put_CustomNumLockText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_CUSTOMNUMLOCKTEXT);

	if(properties.customNumLockText != newValue) {
		properties.customNumLockText.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.panels->ApplySettings(FALSE, TRUE);
		}
		FireOnChanged(DISPID_STATBAR_CUSTOMNUMLOCKTEXT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_CustomScrollLockText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.customScrollLockText.Copy();
	return S_OK;
}

STDMETHODIMP StatusBar::put_CustomScrollLockText(BSTR newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT);

	if(properties.customScrollLockText != newValue) {
		properties.customScrollLockText.AssignBSTR(newValue);
		SetDirty(TRUE);

		if(IsWindow()) {
			properties.panels->ApplySettings(FALSE, TRUE);
		}
		FireOnChanged(DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP StatusBar::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_DISABLEDEVENTS);

	if(properties.disabledEvents != newValue) {
		if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
			if(IsWindow()) {
				if(static_cast<long>(newValue) & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
					panelUnderMouse = -2;
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_STATBAR_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP StatusBar::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_DONTREDRAW);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_STATBAR_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP StatusBar::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_ENABLED);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_STATBAR_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP StatusBar::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_STATBAR_FONT);

	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_STATBAR_FONT);
	return S_OK;
}

STDMETHODIMP StatusBar::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_STATBAR_FONT);

	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_STATBAR_FONT);
	return S_OK;
}

STDMETHODIMP StatusBar::get_ForceSizeGripperDisplay(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.forceSizeGripperDisplay = ((GetStyle() & SBARS_SIZEGRIP) == SBARS_SIZEGRIP);
	}

	*pValue = BOOL2VARIANTBOOL(properties.forceSizeGripperDisplay);
	return S_OK;
}

STDMETHODIMP StatusBar::put_ForceSizeGripperDisplay(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_FORCESIZEGRIPPERDISPLAY);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.forceSizeGripperDisplay != b) {
		properties.forceSizeGripperDisplay = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_STATBAR_FORCESIZEGRIPPERDISPLAY);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP StatusBar::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_HOVERTIME);

	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_STATBAR_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP StatusBar::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP StatusBar::get_MinimumHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.minimumHeight;
	return S_OK;
}

STDMETHODIMP StatusBar::put_MinimumHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_MINIMUMHEIGHT);

	if(properties.minimumHeight != newValue) {
		properties.minimumHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(SB_SETMINHEIGHT, properties.minimumHeight, 0);
			SendMessage(WM_SIZE, 0, 0);
		}
		FireOnChanged(DISPID_STATBAR_MINIMUMHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP StatusBar::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_STATBAR_MOUSEICON);

	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_STATBAR_MOUSEICON);
	return S_OK;
}

STDMETHODIMP StatusBar::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_STATBAR_MOUSEICON);

	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_STATBAR_MOUSEICON);
	return S_OK;
}

STDMETHODIMP StatusBar::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP StatusBar::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_MOUSEPOINTER);

	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_STATBAR_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Panels(IStatusBarPanels** ppPanels)
{
	ATLASSERT_POINTER(ppPanels, IStatusBarPanels*);
	if(!ppPanels) {
		return E_POINTER;
	}

	return properties.panels->QueryInterface(IID_IStatusBarPanels, reinterpret_cast<LPVOID*>(ppPanels));
}

STDMETHODIMP StatusBar::get_PanelToAutoSize(IStatusBarPanel** ppPanel)
{
	ATLASSERT_POINTER(ppPanel, IStatusBarPanel*);
	if(!ppPanel) {
		return E_POINTER;
	}

	*ppPanel = NULL;
	if(properties.panelToAutoSize) {
		return properties.panelToAutoSize->QueryInterface(IID_IStatusBarPanel, reinterpret_cast<LPVOID*>(ppPanel));
	}
	return S_OK;
}

STDMETHODIMP StatusBar::putref_PanelToAutoSize(IStatusBarPanel* pPanel)
{
	PUTPROPPROLOG(DISPID_STATBAR_PANELTOAUTOSIZE);

	if(properties.panelToAutoSize != pPanel) {
		properties.panelToAutoSize = pPanel;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyPanelWidths(FALSE);
		}
		FireOnChanged(DISPID_STATBAR_PANELTOAUTOSIZE);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP StatusBar::get_RegisterForOLEDragDrop(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.registerForOLEDragDrop);
	return S_OK;
}

STDMETHODIMP StatusBar::put_RegisterForOLEDragDrop(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_REGISTERFOROLEDRAGDROP);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.registerForOLEDragDrop != b) {
		properties.registerForOLEDragDrop = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.registerForOLEDragDrop) {
				ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
			} else {
				ATLVERIFY(RevokeDragDrop(*this) == S_OK);
			}
		}
		FireOnChanged(DISPID_STATBAR_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_RightToLeftLayout(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeftLayout = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	}

	*pValue = BOOL2VARIANTBOOL(properties.rightToLeftLayout);
	return S_OK;
}

STDMETHODIMP StatusBar::put_RightToLeftLayout(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_RIGHTTOLEFTLAYOUT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.rightToLeftLayout != b) {
		properties.rightToLeftLayout = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeftLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_STATBAR_RIGHTTOLEFTLAYOUT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_ShowToolTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showToolTips = ((GetStyle() & SBARS_TOOLTIPS) == SBARS_TOOLTIPS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showToolTips);
	return S_OK;
}

STDMETHODIMP StatusBar::put_ShowToolTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_SHOWTOOLTIPS);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showToolTips != b) {
		properties.showToolTips = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_STATBAR_SHOWTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_SimpleMode(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.simpleMode = (SendMessage(SB_ISSIMPLE, 0, 0) != 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.simpleMode);
	return S_OK;
}

STDMETHODIMP StatusBar::put_SimpleMode(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_SIMPLEMODE);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.simpleMode != b) {
		properties.simpleMode = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(SB_SIMPLE, b, 0);
		}
		FireOnChanged(DISPID_STATBAR_SIMPLEMODE);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_SimplePanel(IStatusBarPanel** ppPanel)
{
	ATLASSERT_POINTER(ppPanel, IStatusBarPanel*);
	if(!ppPanel) {
		return E_POINTER;
	}

	return properties.simplePanel->QueryInterface(IID_IStatusBarPanel, reinterpret_cast<LPVOID*>(ppPanel));
}

STDMETHODIMP StatusBar::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP StatusBar::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_SUPPORTOLEDRAGIMAGES);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_STATBAR_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP StatusBar::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP StatusBar::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_STATBAR_USESYSTEMFONT);

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_STATBAR_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP StatusBar::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP StatusBar::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP StatusBar::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP StatusBar::GetBorders(OLE_YSIZE_PIXELS* pHorizontalBorder/* = NULL*/, OLE_XSIZE_PIXELS* pVerticalBorder/* = NULL*/, OLE_XSIZE_PIXELS* pInterPanelBorder/* = NULL*/)
{
	if(IsWindow()) {
		int borders[3];
		ZeroMemory(borders, sizeof(borders));
		if(SendMessage(SB_GETBORDERS, 0, reinterpret_cast<LPARAM>(borders))) {
			if(pHorizontalBorder) {
				*pHorizontalBorder = borders[0];
			}
			if(pVerticalBorder) {
				*pVerticalBorder = borders[1];
			}
			if(pInterPanelBorder) {
				*pInterPanelBorder = borders[2];
			}
		} else {
			return E_FAIL;
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBar::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IStatusBarPanel** ppHitPanel/* = NULL*/)
{
	ATLASSERT_POINTER(ppHitPanel, IStatusBarPanel*);
	if(!ppHitPanel) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT flags = static_cast<UINT>(*pHitTestDetails);
		int panelIndex = HitTest(x, y, &flags);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(flags);
		}
		GetPanelObject(panelIndex, ppHitPanel);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP StatusBar::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBar::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		UpdateWindow();
	}
	return S_OK;
}

STDMETHODIMP StatusBar::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}


LRESULT StatusBar::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT StatusBar::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT StatusBar::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT StatusBar::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL inDesignMode = IsInDesignMode();
	if(!inDesignMode) {
		TCHAR pBuffer[200];
		ZeroMemory(pBuffer, 200 * sizeof(TCHAR));
		GetModuleFileName(NULL, pBuffer, 200);
		PathStripPath(pBuffer);
		if(lstrcmpi(pBuffer, TEXT("vb6.exe")) == 0) {
			HWND hWnd = GetParent();
			while(hWnd) {
				if(GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("DesignerWindow")) == 0) {
					hWnd = ::GetParent(hWnd);
					if(hWnd && GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("MDIClient")) == 0) {
						hWnd = ::GetParent(hWnd);
						if(hWnd && GetClassName(hWnd, pBuffer, 200) && lstrcmpi(pBuffer, TEXT("wndclass_desked_gsk")) == 0) {
							inDesignMode = TRUE;
						}
					}
					break;
				}
				hWnd = ::GetParent(hWnd);
			}
		}
	}
	if(!inDesignMode) {
		DWORD position = GetMessagePos();
		POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		if(WindowFromPoint(mousePosition) == *this) {
			ScreenToClient(&mousePosition);
			SHORT button = 0;
			SHORT shift = 0;
			WPARAM2BUTTONANDSHIFT(-1, button, shift);
			if(HIWORD(lParam) == WM_LBUTTONDOWN) {
				button = 1/*MouseButtonConstants.vbLeftButton*/;
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			} else if(HIWORD(lParam) == WM_MBUTTONDOWN) {
				button = 4/*MouseButtonConstants.vbMiddleButton*/;
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			} else if(HIWORD(lParam) == WM_RBUTTONDOWN) {
				button = 2/*MouseButtonConstants.vbRightButton*/;
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			} else if(HIWORD(lParam) == WM_XBUTTONDOWN) {
				Raise_MouseDown(button, shift, mousePosition.x, mousePosition.y);
			}
			return MA_NOACTIVATEANDEAT;
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT StatusBar::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	return 0;
}

LRESULT StatusBar::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT StatusBar::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT StatusBar::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	WTL::CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT StatusBar::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_STATBAR_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_STATBAR_FONT);
	}

	return lr;
}

LRESULT StatusBar::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT StatusBar::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETNONCLIENTMETRICS) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT StatusBar::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = reinterpret_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT StatusBar::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		case timers.ID_CHECKPANELS:
			UpdatePredefinedPanels();
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT StatusBar::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	WTL::CRect windowRectangle = m_rcPos;
	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull() && pDetails->x != 0 && pDetails->y != 0)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		BOOL raiseEvent = FALSE;
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			ApplyPanelWidths(FALSE);
		}
		if(raiseEvent) {
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT StatusBar::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT StatusBar::OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPDRAWITEMSTRUCT pDetails = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	ATLASSERT(pDetails->hwndItem == *this);
	ATLASSUME(reinterpret_cast<IStatusBarPanel*>(pDetails->itemData));

	IStatusBarPanel* pPanel = reinterpret_cast<IStatusBarPanel*>(pDetails->itemData);
	if(pPanel) {
		PanelContentConstants content = pcText;
		pPanel->get_Content(&content);
		if(content == pcOwnerDrawn) {
			Raise_OwnerDrawPanel(pPanel, HandleToLong(pDetails->hDC), reinterpret_cast<RECTANGLE*>(&pDetails->rcItem));
		} else {
			CComBSTR text;
			pPanel->get_Text(&text);

			if(text.Length() > 0) {
				CDCHandle targetDC = pDetails->hDC;

				LPTSTR pBuffer = NULL;
				if(IsComctl32Version610OrNewer()) {
					pBuffer = StrDup(COLE2CT(text));
				} else {
					// NOTE: Use LocalAlloc here, because this is what StrDup is using.
					pBuffer = reinterpret_cast<LPTSTR>(LocalAlloc(LPTR, (MAX_PANELTEXTLENGTH + 1) * sizeof(TCHAR)));
					if(pBuffer) {
						lstrcpyn(pBuffer, COLE2CT(text), MAX_PANELTEXTLENGTH);
					}
				}

				AlignmentConstants alignment = alLeft;
				pPanel->get_Alignment(&alignment);
				VARIANT_BOOL enabled = VARIANT_TRUE;
				pPanel->get_Enabled(&enabled);
				OLE_HANDLE h = NULL;
				pPanel->get_hIcon(&h);
				HICON hIcon = static_cast<HICON>(LongToHandle(h));
				VARIANT_BOOL rightToLeftText = VARIANT_FALSE;
				pPanel->get_RightToLeftText(&rightToLeftText);

				WTL::CRect textRectangle = pDetails->rcItem;

				UINT numberOfPanels = SendMessage(SB_GETPARTS, 0, 0);
				if(pDetails->itemID == numberOfPanels - 1) {
					CWindow parentWindow = GetParent();
					BOOL hasGripper = properties.forceSizeGripperDisplay || (((parentWindow.GetStyle() & WS_THICKFRAME) == WS_THICKFRAME) && !parentWindow.IsZoomed());
					if(hasGripper) {
						// calculate the gripper's width
						int gripperWidth = 0;

						CTheme themingEngine;
						if(flags.usingThemes) {
							themingEngine.OpenThemeData(*this, VSCLASS_STATUS);
							if(!themingEngine.IsThemeNull()) {
								RECT clientArea = {0};
								GetClientRect(&clientArea);
								SIZE gripperSize = {0};
								ATLVERIFY(SUCCEEDED(themingEngine.GetThemePartSize(targetDC, SP_GRIPPER, 1/*DEF_NORMAL*/, &clientArea, TS_DRAW, &gripperSize)));
								gripperWidth = gripperSize.cx;
							}
						}
						if(gripperWidth == 0) {
							NONCLIENTMETRICS nonClientMetrics = {0};
							nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
							SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
							gripperWidth = nonClientMetrics.iScrollWidth;
						}

						textRectangle.right -= gripperWidth;
					}
				}

				if(hIcon) {
					ICONINFO iconInfo = {0};
					GetIconInfo(hIcon, &iconInfo);
					if(iconInfo.hbmColor) {
						CBitmapHandle bmp = iconInfo.hbmColor;
						WTL::CSize bitmapSize(0, 0);
						bmp.GetSize(bitmapSize);
						textRectangle.left += bitmapSize.cx + GetSystemMetrics(SM_CXEDGE) * 2;
					} else if(iconInfo.hbmMask) {
						CBitmapHandle bmp = iconInfo.hbmMask;
						WTL::CSize bitmapSize(0, 0);
						bmp.GetSize(bitmapSize);
						textRectangle.left += bitmapSize.cx + GetSystemMetrics(SM_CXEDGE) * 2;
					}
					if(iconInfo.hbmColor) {
						DeleteObject(iconInfo.hbmColor);
					}
					if(iconInfo.hbmMask) {
						DeleteObject(iconInfo.hbmMask);
					}
				}
				textRectangle.InflateRect(-1, 0);

				CTheme themingEngine;
				if(flags.usingThemes) {
					/* HACK: The theming engine doesn't seem to offer an "official" way for drawing disabled text,
					         so we misuse the Button theme which defines a state which makes the text look exactly
					         like we want it to look. */
					themingEngine.OpenThemeData(*this, (enabled == VARIANT_FALSE ? VSCLASS_BUTTON : VSCLASS_STATUS));
				}
				if(themingEngine.IsThemeNull()) {
					// don't use any theme
					DWORD flags = DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | (rightToLeftText == VARIANT_FALSE ? 0 : DT_RTLREADING);
					switch(alignment) {
						case alLeft:
							flags |= DT_LEFT;
							break;
						case alCenter:
							flags |= DT_CENTER;
							break;
						case alRight:
							flags |= DT_RIGHT;
							break;
					}

					targetDC.SetBkMode(TRANSPARENT);
					if(enabled == VARIANT_FALSE) {
						WTL::CRect offsetTextRectangle = textRectangle;
						offsetTextRectangle.OffsetRect(1, 1);
						COLORREF previousColor = targetDC.SetTextColor(GetSysColor(COLOR_3DHIGHLIGHT));
						if(pBuffer) {
							targetDC.DrawText(pBuffer, -1, &offsetTextRectangle, flags);
							targetDC.SetTextColor(GetSysColor(COLOR_3DSHADOW));
							targetDC.DrawText(pBuffer, -1, &textRectangle, flags);
						}
						targetDC.SetTextColor(previousColor);
					} else {
						OLE_COLOR foreColor = CLR_NONE;
						pPanel->get_ForeColor(&foreColor);

						if(foreColor == CLR_NONE) {
							if(pBuffer) {
								targetDC.DrawText(pBuffer, -1, &textRectangle, flags);
							}
						} else {
							if(pBuffer) {
								COLORREF previousColor = targetDC.SetTextColor(OLECOLOR2COLORREF(foreColor));
								targetDC.DrawText(pBuffer, -1, &textRectangle, flags);
								targetDC.SetTextColor(previousColor);
							}
						}
					}
				} else {
					// use the current theme
					DWORD flags = DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX | (rightToLeftText == VARIANT_FALSE ? 0 : DT_RTLREADING);
					switch(alignment) {
						case alLeft:
							flags |= DT_LEFT;
							break;
						case alCenter:
							flags |= DT_CENTER;
							break;
						case alRight:
							flags |= DT_RIGHT;
							break;
					}
					if(enabled == VARIANT_FALSE) {
						if(pBuffer) {
							themingEngine.DrawThemeText(targetDC, BP_PUSHBUTTON, PBS_DISABLED, CT2W(pBuffer), -1, flags, DTT_GRAYED, &textRectangle);
						}
					} else {
						OLE_COLOR foreColor = CLR_NONE;
						pPanel->get_ForeColor(&foreColor);

						if(foreColor == CLR_NONE) {
							if(pBuffer) {
								themingEngine.DrawThemeText(targetDC, SP_PANE, 1/*DEF_NORMAL*/, CT2W(pBuffer), -1, flags, 0, &textRectangle);
							}
						} else {
							targetDC.SetBkMode(TRANSPARENT);
							if(pBuffer) {
								COLORREF previousColor = targetDC.SetTextColor(OLECOLOR2COLORREF(foreColor));
								targetDC.DrawText(pBuffer, -1, &textRectangle, flags);
								targetDC.SetTextColor(previousColor);
							}
						}
					}
					themingEngine.CloseThemeData();
				}

				if(pBuffer) {
					LocalFree(pBuffer);
				}
			}

			text.Empty();
		}

		return TRUE;
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT StatusBar::OnSetMinHeight(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_STATBAR_MINIMUMHEIGHT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	properties.minimumHeight = static_cast<OLE_YSIZE_PIXELS>(wParam);
	SetDirty(TRUE);
	FireOnChanged(DISPID_STATBAR_MINIMUMHEIGHT);
	return lr;
}


LRESULT StatusBar::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(CHAR));

	if(pDetails->hdr.idFrom == SB_SIMPLEID) {
		BSTR tooltip = NULL;
		properties.simplePanel->get_ToolTipText(&tooltip);
		if(lstrlenA(CW2A(tooltip)) > 0) {
			lstrcpynA(reinterpret_cast<LPSTR>(pToolTipBuffer), CW2A(tooltip), MAX_PATH);
		}
	} else {
		GuardedSendMessage(*this, SB_GETTIPTEXTA, MAKEWPARAM(pDetails->hdr.idFrom, MAX_PATH), reinterpret_cast<LPARAM>(pToolTipBuffer), NULL);
	}
	pDetails->lpszText = reinterpret_cast<LPSTR>(pToolTipBuffer);

	if(!(properties.disabledEvents & deSetupToolTipWindow)) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(pDetails->hdr.idFrom);
		Raise_SetupToolTipWindow(pPanel, HandleToLong(pDetails->hdr.hwndFrom));
	}

	return 0;
}

LRESULT StatusBar::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	ZeroMemory(pToolTipBuffer, MAX_PATH * sizeof(WCHAR));

	if(pDetails->hdr.idFrom == SB_SIMPLEID) {
		BSTR tooltip = NULL;
		properties.simplePanel->get_ToolTipText(&tooltip);
		if(lstrlenW(CW2CW(tooltip)) > 0) {
			lstrcpynW(reinterpret_cast<LPWSTR>(pToolTipBuffer), CW2CW(tooltip), MAX_PATH);
		}
	} else {
		GuardedSendMessage(*this, SB_GETTIPTEXTW, MAKEWPARAM(pDetails->hdr.idFrom, MAX_PATH), reinterpret_cast<LPARAM>(pToolTipBuffer), NULL);
	}
	pDetails->lpszText = reinterpret_cast<LPWSTR>(pToolTipBuffer);

	if(!(properties.disabledEvents & deSetupToolTipWindow)) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(pDetails->hdr.idFrom);
		Raise_SetupToolTipWindow(pPanel, HandleToLong(pDetails->hdr.hwndFrom));
	}

	return 0;
}

LRESULT StatusBar::OnSimpleModeChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	if(flags.internalSBSIMPLE == 0) {
		Raise_ToggledSimpleMode();

		if(!(properties.disabledEvents & deMouseEvents)) {
			// maybe we have a new panel below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			UINT hitTestDetails = 0;
			int newPanelUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newPanelUnderMouse != panelUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(panelUnderMouse != -2) {
					CComPtr<IStatusBarPanel> pPanelItf = GetPanelObject(panelUnderMouse);
					Raise_PanelMouseLeave(pPanelItf, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
				}
				panelUnderMouse = newPanelUnderMouse;
				if(panelUnderMouse != -2) {
					CComPtr<IStatusBarPanel> pPanelItf = GetPanelObject(panelUnderMouse);
					Raise_PanelMouseEnter(pPanelItf, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
				}
			}
		}
	}

	return 0;
}


void StatusBar::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// retrieve the system font
			NONCLIENTMETRICS nonClientMetrics = {0};
			nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
			properties.font.currentFont.CreateFontIndirect(&nonClientMetrics.lfStatusFont);
			properties.font.owningFontResource = TRUE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update the treeview's font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT StatusBar::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedPanel = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(mouseStatus.lastClickedPanel);
		return Fire_Click(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		/*if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
		}*/

		UINT hitTestDetails = 0;
		int panelIndex = HitTest(x, y, &hitTestDetails);

		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		return Fire_ContextMenu(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	if(panelIndex != mouseStatus.lastClickedPanel) {
		panelIndex = -1;
	}
	mouseStatus.lastClickedPanel = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		return Fire_DblClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	KillTimer(timers.ID_CHECKPANELS);

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedPanel = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(mouseStatus.lastClickedPanel);
		return Fire_MClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	if(panelIndex != mouseStatus.lastClickedPanel) {
		panelIndex = -1;
	}
	mouseStatus.lastClickedPanel = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		return Fire_MDblClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	BOOL fireMouseDown = FALSE;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		if(!mouseStatus.enteredControl) {
			Raise_MouseEnter(button, shift, x, y);
		}
		if(!mouseStatus.hoveredControl) {
			TRACKMOUSEEVENT trackingOptions = {0};
			trackingOptions.cbSize = sizeof(trackingOptions);
			trackingOptions.hwndTrack = *this;
			trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
			TrackMouseEvent(&trackingOptions);

			Raise_MouseHover(button, shift, x, y);
		}
		fireMouseDown = TRUE;
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		if(!mouseStatus.hoveredControl) {
			mouseStatus.HoverControl();
		}
	}
	mouseStatus.StoreClickCandidate(button);
	SetCapture();

	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	POINT pt = mouseStatus.clickPosition;
	ScreenToClient(&pt);
	int clickPanelIndex = HitTest(pt.x, pt.y, static_cast<LPUINT>(NULL));
	if(mouseStatus.IsDoubleClickCandidate(button) && panelIndex == clickPanelIndex) {
		/* Watch for double-clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the double-click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_DblClick(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RDblClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MDblClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XDblClick(button, shift, x, y);
					}
					break;
			}
			mouseStatus.RemoveAllDoubleClickCandidates();
		}

		mouseStatus.RemoveClickCandidate(button);
		if(GetCapture() == *this) {
			ReleaseCapture();
		}
		fireMouseDown = FALSE;
	} else {
		mouseStatus.RemoveAllDoubleClickCandidates();
	}

	HRESULT hr = S_OK;
	if(mouseStatus.IsClickCandidate(button)) {
		mouseStatus.mouseDownPanel = panelIndex;
	}
	if(fireMouseDown) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		hr = Fire_MouseDown(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return hr;
}

inline HRESULT StatusBar::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	panelUnderMouse = panelIndex;

	CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_MouseEnter(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	if(pPanel) {
		Raise_PanelMouseEnter(pPanel, button, shift, x, y, hitTestDetails);
	}
	return hr;
}

inline HRESULT StatusBar::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	mouseStatus.HoverControl();

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		int panelIndex = HitTest(x, y, &hitTestDetails);

		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		return Fire_MouseHover(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelUnderMouse);
	if(pPanel) {
		Raise_PanelMouseLeave(pPanel, button, shift, x, y, hitTestDetails);
	}
	panelUnderMouse = -2;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pPanel = GetPanelObject(panelIndex);
		return Fire_MouseLeave(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
	// Do we move over another panel than before?
	if(panelIndex != panelUnderMouse) {
		CComPtr<IStatusBarPanel> pPrevPanel = GetPanelObject(panelUnderMouse);
		if(pPrevPanel) {
			Raise_PanelMouseLeave(pPrevPanel, button, shift, x, y, hitTestDetails);
		}
		panelUnderMouse = panelIndex;
		if(pPanel) {
			Raise_PanelMouseEnter(pPanel, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);

	HRESULT hr = S_OK;
	if(/*(m_nFreezeEvents == 0) && */!(properties.disabledEvents & deMouseEvents)) {
		hr = Fire_MouseUp(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}

	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea = {0};
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}
		if(GetCapture() == *this) {
			ReleaseCapture();
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents) && (panelIndex == mouseStatus.mouseDownPanel)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents) && (panelIndex == mouseStatus.mouseDownPanel)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents) && (panelIndex == mouseStatus.mouseDownPanel)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
			if(panelIndex == mouseStatus.mouseDownPanel) {
				mouseStatus.StoreDoubleClickCandidate(button);
			}
		}

		mouseStatus.mouseDownPanel = -1;
		mouseStatus.RemoveClickCandidate(button);
	}

	return hr;
}

inline HRESULT StatusBar::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

		UINT hitTestDetails = 0;
		int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
		CComPtr<IStatusBarPanel> pDropTarget = GetPanelObject(dropTarget);

		if(pData) {
			/* Actually we wouldn't need the next line, because the data object passed to this method should
			   always be the same as the data object that was passed to Raise_OLEDragEnter. */
			dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
		} else {
			dragDropStatus.pActiveDataObject = NULL;
		}
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT StatusBar::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			CComPtr<IStatusBarPanel> pDropTarget = GetPanelObject(dropTarget);
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.lastDropTarget = dropTarget;
	return hr;
}

inline HRESULT StatusBar::Raise_OLEDragLeave(void)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);

		UINT hitTestDetails = 0;
		int dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
		CComPtr<IStatusBarPanel> pDropTarget = GetPanelObject(dropTarget);

		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT StatusBar::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			CComPtr<IStatusBarPanel> pDropTarget = GetPanelObject(dropTarget);
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.lastDropTarget = dropTarget;
	return hr;
}

inline HRESULT StatusBar::Raise_OwnerDrawPanel(IStatusBarPanel* pPanel, LONG hDC, RECTANGLE* pDrawingRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OwnerDrawPanel(pPanel, hDC, pDrawingRectangle);
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_PanelMouseEnter(IStatusBarPanel* pPanel, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_PanelMouseEnter(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT StatusBar::Raise_PanelMouseLeave(IStatusBarPanel* pPanel, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_PanelMouseLeave(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT StatusBar::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedPanel = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(mouseStatus.lastClickedPanel);
		return Fire_RClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	if(panelIndex != mouseStatus.lastClickedPanel) {
		panelIndex = -1;
	}
	mouseStatus.lastClickedPanel = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		return Fire_RDblClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}
	SetTimer(timers.ID_CHECKPANELS, timers.INT_CHECKPANELS);

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_SetupToolTipWindow(IStatusBarPanel* pPanel, LONG hWndToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SetupToolTipWindow(pPanel, hWndToolTip);
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_ToggledSimpleMode(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ToggledSimpleMode();
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedPanel = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(mouseStatus.lastClickedPanel);
		return Fire_XClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT StatusBar::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int panelIndex = HitTest(x, y, &hitTestDetails);
	if(panelIndex != mouseStatus.lastClickedPanel) {
		panelIndex = -1;
	}
	mouseStatus.lastClickedPanel = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IStatusBarPanel> pPanel = GetPanelObject(panelIndex);
		return Fire_XDblClick(pPanel, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void StatusBar::EnterInternalSBSIMPLESection(void)
{
	++flags.internalSBSIMPLE;
}

void StatusBar::LeaveInternalSBSIMPLESection(void)
{
	--flags.internalSBSIMPLE;
}


void StatusBar::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

DWORD StatusBar::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeftLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	return extendedStyle;
}

DWORD StatusBar::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	/* We can't detect changes to the Align property, i. e. we can't change the window style dynamically.
	   As a consequence we have to use window styles that work with all alignments, so always set
	   CCS_NORESIZE (it doesn't seem to make trouble with bottom-alignment). */
	style |= CCS_NORESIZE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.forceSizeGripperDisplay) {
		style |= SBARS_SIZEGRIP;
	}
	if(properties.showToolTips) {
		style |= SBARS_TOOLTIPS;
	}

	if(m_spClientSite) {
		IOleControlSite* pControlSite = NULL;
		m_spClientSite->QueryInterface(IID_IOleControlSite, reinterpret_cast<LPVOID*>(&pControlSite));
		if(pControlSite) {
			IDispatch* pExtendedControl = NULL;
			pControlSite->GetExtendedControl(&pExtendedControl);
			if(pExtendedControl) {
				LPWSTR pPropertyName = OLESTR("Align");
				DISPID dispidAlign = 0;
				if(pExtendedControl->GetIDsOfNames(IID_NULL, &pPropertyName, 1, LOCALE_SYSTEM_DEFAULT, &dispidAlign) == S_OK) {
					VARIANT v;
					VariantInit(&v);
					DISPPARAMS params = {NULL, NULL, 0, 0};
					if(SUCCEEDED(pExtendedControl->Invoke(dispidAlign, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &v, NULL, NULL))) {
						if(SUCCEEDED(VariantChangeType(&v, &v, 0, VT_I4))) {
							if(v.lVal != 2/*vbAlignBottom*/) {
								style |= CCS_NOPARENTALIGN;
							}
						}
					}
					VariantClear(&v);
				}
				pExtendedControl->Release();
			}
			pControlSite->Release();
		}
	}
	return style;
}

void StatusBar::ApplyPanelWidths(BOOL completeUpdate)
{
	int numberOfPanels = 0;

	LONG l = 0;
	if(SUCCEEDED(properties.panels->Count(&l))) {
		numberOfPanels = l;
	}

	if(numberOfPanels <= 0) {
		if(IsWindow()) {
			// there's always one panel, so make it consume the whole space
			int i = -1;
			SendMessage(SB_SETPARTS, 1, reinterpret_cast<LPARAM>(&i));
		}
		return;
	}

	int panelToAutoSize = -1;
	int autoSizePanelWidth = 0;
	if(properties.panelToAutoSize) {
		l = -1;
		properties.panelToAutoSize->get_Index(&l);
		panelToAutoSize = l;

		if(panelToAutoSize == numberOfPanels - 1) {
			// let msctls_statusbar32 do the dirty work
			autoSizePanelWidth = -1;
		} else if(panelToAutoSize < numberOfPanels) {
			// calculate the remaining space
			// TODO: Is this correct or do we have to consider borders and so on?
			RECT clientRect = {0};
			GetClientRect(&clientRect);
			int remainingWidth = clientRect.right - clientRect.left;
			int minimumWidth = 0;
			for(int panelIndex = 0; panelIndex < numberOfPanels; ++panelIndex) {
				IStatusBarPanel* pPanel = NULL;
				if(SUCCEEDED(properties.panels->get_Item(panelIndex, &pPanel)) && pPanel) {
					if(panelIndex == panelToAutoSize) {
						if(SUCCEEDED(pPanel->get_MinimumWidth(&l))) {
							minimumWidth = l;
						}
					} else {
						if(SUCCEEDED(pPanel->get_PreferredWidth(&l))) {
							remainingWidth -= l;
						}
					}
					pPanel->Release();
				}
			}
			if(remainingWidth < 0) {
				remainingWidth = 0;
			}
			autoSizePanelWidth = max(remainingWidth, minimumWidth);
		}
	}

	PINT pPanelWidths = new int[numberOfPanels];
	int incrementalWidth = 0;
	for(int panelIndex = 0; panelIndex < numberOfPanels; ++panelIndex) {
		IStatusBarPanel* pPanel = NULL;
		if(panelIndex == panelToAutoSize) {
			if(autoSizePanelWidth == -1) {
				incrementalWidth = -1;
			} else {
				incrementalWidth += autoSizePanelWidth;
			}
		} else {
			if(SUCCEEDED(properties.panels->get_Item(panelIndex, &pPanel)) && pPanel) {
				if(SUCCEEDED(pPanel->get_PreferredWidth(&l))) {
					incrementalWidth += l;
				}
				pPanel->Release();
			}
		}
		pPanelWidths[panelIndex] = incrementalWidth;
	}

	if(IsWindow()) {
		SendMessage(SB_SETPARTS, numberOfPanels, reinterpret_cast<LPARAM>(pPanelWidths));
	}
	delete[] pPanelWidths;

	if(completeUpdate) {
		properties.panels->ApplySettings(FALSE, TRUE);
	}
}

void StatusBar::UpdatePredefinedPanels(void)
{
	// HACK: unfortunately we don't receive any WM_INPUTLANGCHANGE message, so use the following ugly hack
	static HKL currentKeyboardLayout = GetKeyboardLayout(0);
	if(currentKeyboardLayout != GetKeyboardLayout(0)) {
		currentKeyboardLayout = GetKeyboardLayout(0);
		properties.panels->ApplySettings(FALSE, TRUE);
		return;
	}

	int numberOfPanels = 0;
	LONG l = 0;
	if(SUCCEEDED(properties.panels->Count(&l))) {
		numberOfPanels = l;
	}

	if(numberOfPanels <= 0) {
		// there's nothing to do
		return;
	}

	for(int panelIndex = 0; panelIndex < numberOfPanels; ++panelIndex) {
		IStatusBarPanel* pPanel = NULL;
		if(SUCCEEDED(properties.panels->get_Item(panelIndex, &pPanel)) && pPanel) {
			PanelContentConstants content = pcText;
			pPanel->get_Content(&content);

			switch(content) {
				case pcCapsLock:
					pPanel->put_Enabled(BOOL2VARIANTBOOL(GetKeyState(VK_CAPITAL) & 0x1));
					break;
				case pcNumLock:
					pPanel->put_Enabled(BOOL2VARIANTBOOL(GetKeyState(VK_NUMLOCK) & 0x1));
					break;
				case pcScrollLock:
					pPanel->put_Enabled(BOOL2VARIANTBOOL(GetKeyState(VK_SCROLL) & 0x1));
					break;
				case pcKanaLock:
					pPanel->put_Enabled(BOOL2VARIANTBOOL(GetKeyState(VK_KANA) & 0x1));
					break;
				case pcInsertKey: {
					BYTE keyboardState[256];
					GetKeyboardState(keyboardState);
					pPanel->put_Enabled(BOOL2VARIANTBOOL(keyboardState[VK_INSERT] & 0x1));
					break;
				}
				case pcShortTime: {
					TCHAR buffer[81];
					GetTimeFormat(GetThreadLocale(), TIME_NOSECONDS, NULL, NULL, buffer, 80);
					pPanel->put_Text(_bstr_t(buffer).Detach());
					break;
				}
				case pcLongTime: {
					TCHAR buffer[81];
					GetTimeFormat(GetThreadLocale(), 0, NULL, NULL, buffer, 80);
					pPanel->put_Text(_bstr_t(buffer).Detach());
					break;
				}
				case pcShortDate: {
					TCHAR buffer[81];
					GetDateFormat(GetThreadLocale(), DATE_SHORTDATE, NULL, NULL, buffer, 80);
					pPanel->put_Text(_bstr_t(buffer).Detach());
					break;
				}
				case pcLongDate: {
					TCHAR buffer[81];
					GetDateFormat(GetThreadLocale(), DATE_LONGDATE, NULL, NULL, buffer, 80);
					pPanel->put_Text(_bstr_t(buffer).Detach());
					break;
				}
			}

			pPanel->Release();
		}
	}
}

CComPtr<IStatusBarPanel> StatusBar::GetPanelObject(int panelIndex)
{
	CComPtr<IStatusBarPanel> pPanel = NULL;
	GetPanelObject(panelIndex, &pPanel);
	return pPanel;
}

void StatusBar::GetPanelObject(int panelIndex, IStatusBarPanel** ppPanel)
{
	ATLASSERT_POINTER(ppPanel, IStatusBarPanel*);
	ATLASSUME(ppPanel);

	*ppPanel = NULL;
	if(panelIndex == -1) {
		properties.simplePanel->QueryInterface(IID_IStatusBarPanel, reinterpret_cast<LPVOID*>(ppPanel));
	} else {
		properties.panels->get_Item(panelIndex, ppPanel);
	}
}

void StatusBar::SendConfigurationMessages(void)
{
	SendMessage(SB_SETBKCOLOR, 0, (properties.backColor == CLR_NONE ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.backColor)));
	SendMessage(SB_SETMINHEIGHT, properties.minimumHeight, 0);

	EnterInternalSBSIMPLESection();
	SendMessage(SB_SIMPLE, TRUE, 0);
	properties.simplePanel->ApplySettings(FALSE, TRUE);
	SendMessage(SB_SIMPLE, FALSE, 0);
	ApplyPanelWidths(TRUE);
	LeaveInternalSBSIMPLESection();

	if(properties.simpleMode) {
		SendMessage(SB_SIMPLE, TRUE, 0);
	} else {
		Raise_ToggledSimpleMode();
	}

	ApplyFont();
}

HCURSOR StatusBar::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

int StatusBar::HitTest(LONG x, LONG y, LPUINT pFlags)
{
	ATLASSERT(IsWindow());

	UINT flags = 0;
	int panelIndex = -2;
	RECT clientRectangle = {0};
	GetClientRect(&clientRectangle);

	if(x < clientRectangle.left) {
		flags |= SBARHT_TOLEFT;
	} else if(x > clientRectangle.right) {
		flags |= SBARHT_TORIGHT;
	}
	if(y < clientRectangle.top) {
		flags |= SBARHT_ABOVE;
	} else if(y > clientRectangle.bottom) {
		flags |= SBARHT_BELOW;
	}

	if(flags == 0) {
		if(SendMessage(SB_ISSIMPLE, 0, 0)) {
			panelIndex = -1;
			flags = SBARHT_ONITEM;
		} else {
			flags = SBARHT_NOWHERE;
			int panels = SendMessage(SB_GETPARTS, 0, 0);
			for(int i = 0; i < panels; ++i) {
				RECT boundingRectangle = {0};
				if(SendMessage(SB_GETRECT, i, reinterpret_cast<LPARAM>(&boundingRectangle))) {
					if(x < boundingRectangle.right) {
						flags = SBARHT_ONITEM;
						panelIndex = i;
						break;
					}
				}
			}
		}
	}

	if(pFlags) {
		*pFlags = flags;
	}
	return panelIndex;
}

BOOL StatusBar::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}


BOOL StatusBar::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}