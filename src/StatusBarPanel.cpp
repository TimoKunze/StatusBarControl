// StatusBarPanel.cpp: A wrapper for existing status bar panels.

#include "stdafx.h"
#include "StatusBarPanel.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP StatusBarPanel::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IStatusBarPanel, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP StatusBarPanel::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	ATLASSUME(pPropertyBag);
	if(!pPropertyBag) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	if(SUCCEEDED(pPropertyBag->Read(OLESTR("Version"), &propertyValue, pErrorLog))) {
		if(propertyValue.lVal > 0x0102) {
			return E_FAIL;
		}
	}
	LONG version = propertyValue.lVal;

	AlignmentConstants alignment = alLeft;
	if(properties.panelIndex != SB_SIMPLEID) {
		hr = pPropertyBag->Read(OLESTR("Alignment"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		alignment = static_cast<AlignmentConstants>(propertyValue.lVal);
	}
	hr = pPropertyBag->Read(OLESTR("BorderStyle"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	PanelBorderStyleConstants borderStyle = static_cast<PanelBorderStyleConstants>(propertyValue.lVal);
	PanelContentConstants content = pcText;
	VARIANT_BOOL enabled = VARIANT_TRUE;
	OLE_COLOR foreColor = CLR_NONE;
	LONG minimumWidth = 100;
	if(properties.panelIndex != SB_SIMPLEID) {
		hr = pPropertyBag->Read(OLESTR("Content"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		content = static_cast<PanelContentConstants>(propertyValue.lVal);
		propertyValue.vt = VT_BOOL;
		hr = pPropertyBag->Read(OLESTR("Enabled"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		enabled = propertyValue.boolVal;
		propertyValue.vt = VT_I4;
		hr = pPropertyBag->Read(OLESTR("ForeColor"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		foreColor = propertyValue.lVal;
		hr = pPropertyBag->Read(OLESTR("MinimumWidth"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		minimumWidth = propertyValue.lVal;
	}
	hr = pPropertyBag->Read(OLESTR("PanelData"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	LONG panelData = propertyValue.lVal;
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("ParseTabs"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL parseTabs = propertyValue.boolVal;
	LONG preferredWidth = 100;
	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.vt = VT_I4;
		hr = pPropertyBag->Read(OLESTR("PreferredWidth"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		preferredWidth = propertyValue.lVal;
	}
	propertyValue.vt = VT_BOOL;
	hr = pPropertyBag->Read(OLESTR("RightToLeftText"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	VARIANT_BOOL rightToLeftText = propertyValue.boolVal;
	CComBSTR text;
	CComBSTR toolTipText;
	if(version < 0x0102)
	{
		propertyValue.vt = VT_BSTR;
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		text = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
		propertyValue.vt = VT_BSTR;
		hr = pPropertyBag->Read(OLESTR("ToolTipText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		toolTipText = propertyValue.bstrVal;
		SysFreeString(propertyValue.bstrVal);
	} else {
		propertyValue.vt = VT_ARRAY | VT_UI1;
		hr = pPropertyBag->Read(OLESTR("Text"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			text.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
		hr = pPropertyBag->Read(OLESTR("ToolTipText"), &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		if(propertyValue.vt == (VT_ARRAY | VT_UI1) && propertyValue.parray) {
			toolTipText.ArrayToBSTR(propertyValue.parray);
			VariantClear(&propertyValue);
		} else {
			VariantClear(&propertyValue);
			return E_FAIL;
		}
	}

	if(properties.panelIndex != SB_SIMPLEID) {
		hr = put_Alignment(alignment);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		hr = put_Content(content);
		if(FAILED(hr)) {
			return hr;
		}
		hr = put_Enabled(enabled);
		if(FAILED(hr)) {
			return hr;
		}
		hr = put_ForeColor(foreColor);
		if(FAILED(hr)) {
			return hr;
		}
		hr = put_MinimumWidth(minimumWidth);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_PanelData(panelData);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ParseTabs(parseTabs);
	if(FAILED(hr)) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		hr = put_PreferredWidth(preferredWidth);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_RightToLeftText(rightToLeftText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ToolTipText(toolTipText);
	if(FAILED(hr)) {
		return hr;
	}

	m_bRequiresSave = FALSE;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/)
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

	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.lVal = properties.alignment;
		if(FAILED(hr = pPropertyBag->Write(OLESTR("Alignment"), &propertyValue))) {
			return hr;
		}
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("BorderStyle"), &propertyValue))) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.lVal = properties.content;
		if(FAILED(hr = pPropertyBag->Write(OLESTR("Content"), &propertyValue))) {
			return hr;
		}
		propertyValue.vt = VT_BOOL;
		propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
		if(FAILED(hr = pPropertyBag->Write(OLESTR("Enabled"), &propertyValue))) {
			return hr;
		}
		propertyValue.vt = VT_I4;
		propertyValue.lVal = properties.foreColor;
		if(FAILED(hr = pPropertyBag->Write(OLESTR("ForeColor"), &propertyValue))) {
			return hr;
		}
		propertyValue.lVal = properties.minimumWidth;
		if(FAILED(hr = pPropertyBag->Write(OLESTR("MinimumWidth"), &propertyValue))) {
			return hr;
		}
	}
	propertyValue.lVal = properties.panelData;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("PanelData"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.parseTabs);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ParseTabs"), &propertyValue))) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.vt = VT_I4;
		propertyValue.lVal = properties.preferredWidth;
		if(FAILED(hr = pPropertyBag->Write(OLESTR("PreferredWidth"), &propertyValue))) {
			return hr;
		}
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftText);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("RightToLeftText"), &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.text.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Text"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);
	propertyValue.vt = VT_ARRAY | VT_UI1;
	properties.toolTipText.BSTRToArray(&propertyValue.parray);
	if(FAILED(hr = pPropertyBag->Write(OLESTR("ToolTipText"), &propertyValue))) {
		VariantClear(&propertyValue);
		return hr;
	}
	VariantClear(&propertyValue);

	if(clearDirtyFlag) {
		m_bRequiresSave = FALSE;
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/);

	if(properties.panelIndex == SB_SIMPLEID) {
		// we've 2 VT_I4 properties...
		pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(LONG));
		// ...and 2 VT_BOOL properties...
		pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	} else {
		// we've 7 VT_I4 properties...
		pSize->QuadPart += 7 * (sizeof(VARTYPE) + sizeof(LONG));
		// ...and 3 VT_BOOL properties...
		pSize->QuadPart += 3 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));
	}
	// ...and 2 VT_BSTR properties...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.text.ByteLength() + sizeof(OLECHAR);
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(ULONG) + properties.toolTipText.ByteLength() + sizeof(OLECHAR);

	return S_OK;
}

STDMETHODIMP StatusBarPanel::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
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

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = NULL;
	if(version == 0x0100) {
		pfnReadVariantFromStream = ReadVariantFromStream_Legacy;
	} else {
		pfnReadVariantFromStream = ReadVariantFromStream;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	AlignmentConstants alignment = alLeft;
	if(properties.panelIndex != SB_SIMPLEID) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		alignment = static_cast<AlignmentConstants>(propertyValue.lVal);
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	PanelBorderStyleConstants borderStyle = static_cast<PanelBorderStyleConstants>(propertyValue.lVal);
	PanelContentConstants content = pcText;
	VARIANT_BOOL enabled = VARIANT_TRUE;
	OLE_COLOR foreColor = CLR_NONE;
	LONG minimumWidth = 100;
	if(properties.panelIndex != SB_SIMPLEID) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		content = static_cast<PanelContentConstants>(propertyValue.lVal);
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
			return hr;
		}
		enabled = propertyValue.boolVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		foreColor = propertyValue.lVal;
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		minimumWidth = propertyValue.lVal;
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG panelData = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL parseTabs = propertyValue.boolVal;
	LONG preferredWidth = 100;
	if(properties.panelIndex != SB_SIMPLEID) {
		if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
			return hr;
		}
		preferredWidth = propertyValue.lVal;
	}
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL rightToLeftText = propertyValue.boolVal;
	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR text;
	if(FAILED(hr = text.ReadFromStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_BSTR)) {
		return hr;
	}
	CComBSTR toolTipText;
	if(FAILED(hr = toolTipText.ReadFromStream(pStream))) {
		return hr;
	}

	if(properties.panelIndex != SB_SIMPLEID) {
		hr = put_Alignment(alignment);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		hr = put_Content(content);
		if(FAILED(hr)) {
			return hr;
		}
		hr = put_Enabled(enabled);
		if(FAILED(hr)) {
			return hr;
		}
		hr = put_ForeColor(foreColor);
		if(FAILED(hr)) {
			return hr;
		}
		hr = put_MinimumWidth(minimumWidth);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_PanelData(panelData);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ParseTabs(parseTabs);
	if(FAILED(hr)) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		hr = put_PreferredWidth(preferredWidth);
		if(FAILED(hr)) {
			return hr;
		}
	}
	hr = put_RightToLeftText(rightToLeftText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Text(text);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ToolTipText(toolTipText);
	if(FAILED(hr)) {
		return hr;
	}

	m_bRequiresSave = FALSE;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG version = 0x0101;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.lVal = properties.alignment;
		if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
			return hr;
		}
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.lVal = properties.content;
		if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
			return hr;
		}
		propertyValue.vt = VT_BOOL;
		propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
		if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
			return hr;
		}
		propertyValue.vt = VT_I4;
		propertyValue.lVal = properties.foreColor;
		if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
			return hr;
		}
		propertyValue.lVal = properties.minimumWidth;
		if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
			return hr;
		}
	}
	propertyValue.lVal = properties.panelData;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.parseTabs);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	if(properties.panelIndex != SB_SIMPLEID) {
		propertyValue.vt = VT_I4;
		propertyValue.lVal = properties.preferredWidth;
		if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
			return hr;
		}
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.rightToLeftText);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	VARTYPE vt = VT_BSTR;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.text.WriteToStream(pStream))) {
		return hr;
	}
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(FAILED(hr = properties.toolTipText.WriteToStream(pStream))) {
		return hr;
	}

	if(clearDirtyFlag) {
		m_bRequiresSave = FALSE;
	}
	return S_OK;
}


StatusBarPanel::Properties::~Properties()
{
	if(pOwnerStatBar) {
		pOwnerStatBar->Release();
	}
}

HWND StatusBarPanel::Properties::GetStatBarHWnd(void)
{
	ATLASSUME(pOwnerStatBar);

	OLE_HANDLE handle = NULL;
	pOwnerStatBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void StatusBarPanel::Attach(int panelIndex)
{
	properties.panelIndex = panelIndex;
}

void StatusBarPanel::Detach(void)
{
	properties.panelIndex = -1;
}

void StatusBarPanel::ApplySettingsTo(int panelIndex, BOOL textAndStyleOnly)
{
	HWND hWndStatBar = properties.GetStatBarHWnd();
	ATLASSERT(IsWindow(hWndStatBar));

	LPTSTR pBuffer = NULL;
	if(properties.text.Length() > 0) {
		if(properties.pOwnerStatBar->IsComctl32Version610OrNewer()) {
			pBuffer = StrDup(COLE2CT(properties.text));
		} else {
			// NOTE: Use LocalAlloc here, because this is what StrDup is using.
			pBuffer = reinterpret_cast<LPTSTR>(LocalAlloc(LPTR, (MAX_PANELTEXTLENGTH + 1) * sizeof(TCHAR)));
			if(pBuffer) {
				lstrcpyn(pBuffer, COLE2CT(properties.text), MAX_PANELTEXTLENGTH);
			}
		}
	}

	UINT style = 0;
	switch(properties.borderStyle) {
		case pbsNone:
			style |= SBT_NOBORDERS;
			break;
		case pbsRaised:
			style |= SBT_POPOUT;
			break;
	}
	if(panelIndex != SB_SIMPLEID) {
		if(properties.alignment != alLeft) {
			style |= SBT_OWNERDRAW;
		}
		if(properties.content == pcOwnerDrawn) {
			style |= SBT_OWNERDRAW;
		}
		if(!properties.enabled) {
			style |= SBT_OWNERDRAW;
		}
		if(properties.foreColor != CLR_NONE) {
			style |= SBT_OWNERDRAW;
		}
	}
	if(!properties.parseTabs/* && properties.pOwnerStatBar->IsComctl32Version580OrNewer()*/) {
		style |= SBT_NOTABPARSING;
	}
	if(properties.rightToLeftText) {
		style |= SBT_RTLREADING;
	}

	SendMessage(hWndStatBar, SB_SETTEXT, panelIndex | style, (style & SBT_OWNERDRAW ? reinterpret_cast<LPARAM>(static_cast<IStatusBarPanel*>(this)) : reinterpret_cast<LPARAM>(pBuffer)));
	if(pBuffer) {
		LocalFree(pBuffer);
		pBuffer = NULL;
	}

	if(!textAndStyleOnly) {
		TCHAR buffer[MAX_PATH + 1];
		if(properties.toolTipText.Length() > 0) {
			lstrcpyn(buffer, COLE2CT(properties.toolTipText), MAX_PATH);
			pBuffer = buffer;
		}
		SendMessage(hWndStatBar, SB_SETTIPTEXT, panelIndex, reinterpret_cast<LPARAM>(pBuffer));

		if(panelIndex == SB_SIMPLEID) {
			SendMessage(hWndStatBar, SB_SETICON, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(properties.hIcon));
		} else {
			SendMessage(hWndStatBar, SB_SETICON, panelIndex, reinterpret_cast<LPARAM>(properties.hIcon));
		}
	}
}

void StatusBarPanel::ApplySettings(BOOL textAndStyleOnly, BOOL updatePredefinedPanels)
{
	if(updatePredefinedPanels && (properties.panelIndex != SB_SIMPLEID)) {
		// determine the predefined panel's text
		switch(properties.content) {
			case pcCapsLock: {
				BSTR customText = L"";
				properties.pOwnerStatBar->get_CustomCapsLockText(&customText);

				if(lstrlen(COLE2CT(customText)) > 0) {
					// use custom text
					properties.text.Attach(customText);
				} else {
					SysFreeString(customText);

					int scanCode = MapVirtualKeyEx(VK_CAPITAL, 0, GetKeyboardLayout(0));
					// we don't care about left and right
					scanCode |= 0x0200;
					TCHAR buffer[81];
					GetKeyNameText(scanCode << 16, buffer, 80);
					properties.text = buffer;
				}
				properties.enabled = GetKeyState(VK_CAPITAL) & 0x1;
				break;
			}
			case pcNumLock: {
				BSTR customText = L"";
				properties.pOwnerStatBar->get_CustomNumLockText(&customText);

				if(lstrlen(COLE2CT(customText)) > 0) {
					// use custom text
					properties.text.Attach(customText);
				} else {
					SysFreeString(customText);

					int scanCode = MapVirtualKeyEx(VK_NUMLOCK, 0, GetKeyboardLayout(0));
					// this key requires the extended bit to be set, also we don't care about left and right
					scanCode |= 0x0300;
					TCHAR buffer[81];
					GetKeyNameText(scanCode << 16, buffer, 80);
					properties.text = buffer;
				}
				properties.enabled = GetKeyState(VK_NUMLOCK) & 0x1;
				break;
			}
			case pcScrollLock: {
				BSTR customText = L"";
				properties.pOwnerStatBar->get_CustomScrollLockText(&customText);

				if(lstrlen(COLE2CT(customText)) > 0) {
					// use custom text
					properties.text.Attach(customText);
				} else {
					SysFreeString(customText);

					int scanCode = MapVirtualKeyEx(VK_SCROLL, 0, GetKeyboardLayout(0));
					// we don't care about left and right
					scanCode |= 0x0200;
					TCHAR buffer[81];
					GetKeyNameText(scanCode << 16, buffer, 80);
					properties.text = buffer;
				}
				properties.enabled = GetKeyState(VK_SCROLL) & 0x1;
				break;
			}
			case pcKanaLock: {
				BSTR customText = L"";
				properties.pOwnerStatBar->get_CustomKanaLockText(&customText);

				if(lstrlen(COLE2CT(customText)) > 0) {
					// use custom text
					properties.text.Attach(customText);
				} else {
					SysFreeString(customText);

					properties.text = TEXT("KANA");
				}
				properties.enabled = GetKeyState(VK_KANA) & 0x1;
				break;
			}
			case pcInsertKey: {
				BSTR customText = L"";
				properties.pOwnerStatBar->get_CustomInsertKeyText(&customText);

				if(lstrlen(COLE2CT(customText)) > 0) {
					// use custom text
					properties.text.Attach(customText);
				} else {
					SysFreeString(customText);

					int scanCode = MapVirtualKeyEx(VK_INSERT, 0, GetKeyboardLayout(0));
					// this key requires the extended bit to be set, also we don't care about left and right
					scanCode |= 0x0300;
					TCHAR buffer[81];
					GetKeyNameText(scanCode << 16, buffer, 80);
					properties.text = buffer;
				}

				BYTE keyboardState[256];
				GetKeyboardState(keyboardState);
				properties.enabled = keyboardState[VK_INSERT] & 0x1;
				break;
			}
			case pcShortTime: {
				TCHAR buffer[81];
				GetTimeFormat(GetThreadLocale(), TIME_NOSECONDS, NULL, NULL, buffer, 80);
				properties.text = buffer;
				break;
			}
			case pcLongTime: {
				TCHAR buffer[81];
				GetTimeFormat(GetThreadLocale(), 0, NULL, NULL, buffer, 80);
				properties.text = buffer;
				break;
			}
			case pcShortDate: {
				TCHAR buffer[81];
				GetDateFormat(GetThreadLocale(), DATE_SHORTDATE, NULL, NULL, buffer, 80);
				properties.text = buffer;
				break;
			}
			case pcLongDate: {
				TCHAR buffer[81];
				GetDateFormat(GetThreadLocale(), DATE_LONGDATE, NULL, NULL, buffer, 80);
				properties.text = buffer;
				break;
			}
		}
	}

	ApplySettingsTo(properties.panelIndex, textAndStyleOnly);
}

void StatusBarPanel::SetOwner(StatusBar* pOwner)
{
	if(properties.pOwnerStatBar) {
		properties.pOwnerStatBar->Release();
	}
	properties.pOwnerStatBar = pOwner;
	if(properties.pOwnerStatBar) {
		properties.pOwnerStatBar->AddRef();
	}
}


STDMETHODIMP StatusBarPanel::get_Alignment(AlignmentConstants* pValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	ATLASSERT_POINTER(pValue, AlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.alignment;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_Alignment(AlignmentConstants newValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	if(properties.alignment != newValue) {
		properties.alignment = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
			}

			ApplySettings(TRUE, FALSE);

			if(simpleMode) {
				SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_BorderStyle(PanelBorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, PanelBorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_BorderStyle(PanelBorderStyleConstants newValue)
{
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, !simpleMode, 0);
			}

			ApplySettings(TRUE, FALSE);

			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				SendMessage(hWndStatBar, SB_SIMPLE, simpleMode, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}

	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_Content(PanelContentConstants* pValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	ATLASSERT_POINTER(pValue, PanelContentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.content;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_Content(PanelContentConstants newValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	if(properties.content != newValue) {
		properties.content = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
			}

			ApplySettings(TRUE, TRUE);

			if(simpleMode) {
				SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_CurrentWidth(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndStatBar = properties.GetStatBarHWnd();
	ATLASSERT(IsWindow(hWndStatBar));

	if(properties.panelIndex == SB_SIMPLEID) {
		*pValue = -1;
		return S_OK;
	} else {
		int panels = SendMessage(hWndStatBar, SB_GETPARTS, 0, 0);
		ATLASSERT(properties.panelIndex < panels);
		if(panels > 0 && properties.panelIndex < panels) {
			PINT pPanelWidths = new int[panels];
			SendMessage(hWndStatBar, SB_GETPARTS, panels, reinterpret_cast<LPARAM>(pPanelWidths));

			if(properties.panelIndex == 0) {
				*pValue = pPanelWidths[properties.panelIndex];
			} else if(pPanelWidths[properties.panelIndex] == -1) {
				*pValue = -1;
			} else {
				*pValue = pPanelWidths[properties.panelIndex] - pPanelWidths[properties.panelIndex - 1];
			}

			delete[] pPanelWidths;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP StatusBarPanel::get_Enabled(VARIANT_BOOL* pValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_Enabled(VARIANT_BOOL newValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;

		switch(properties.content) {
			case pcCapsLock: {
				BOOL isActive = GetKeyState(VK_CAPITAL) & 0x1;
				if(isActive != static_cast<BOOL>(properties.enabled)) {
					// toggle VK_CAPITAL
					INPUT keyboardInput[2];
					ZeroMemory(keyboardInput, sizeof(keyboardInput));
					keyboardInput[0].type = INPUT_KEYBOARD;
					keyboardInput[0].ki.wVk = VK_CAPITAL;
					keyboardInput[1] = keyboardInput[0];
					keyboardInput[1].ki.dwFlags = KEYEVENTF_KEYUP;
					SendInput(2, keyboardInput, sizeof(INPUT));
				}
				break;
			}
			case pcNumLock: {
				BOOL isActive = GetKeyState(VK_NUMLOCK) & 0x1;
				if(isActive != static_cast<BOOL>(properties.enabled)) {
					// toggle VK_NUMLOCK
					INPUT keyboardInput[2];
					ZeroMemory(keyboardInput, sizeof(keyboardInput));
					keyboardInput[0].type = INPUT_KEYBOARD;
					keyboardInput[0].ki.wVk = VK_NUMLOCK;
					keyboardInput[1] = keyboardInput[0];
					keyboardInput[1].ki.dwFlags = KEYEVENTF_KEYUP;
					SendInput(2, keyboardInput, sizeof(INPUT));
				}
				break;
			}
			case pcScrollLock: {
				BOOL isActive = GetKeyState(VK_SCROLL) & 0x1;
				if(isActive != static_cast<BOOL>(properties.enabled)) {
					// toggle VK_SCROLL
					INPUT keyboardInput[2];
					ZeroMemory(keyboardInput, sizeof(keyboardInput));
					keyboardInput[0].type = INPUT_KEYBOARD;
					keyboardInput[0].ki.wVk = VK_SCROLL;
					keyboardInput[1] = keyboardInput[0];
					keyboardInput[1].ki.dwFlags = KEYEVENTF_KEYUP;
					SendInput(2, keyboardInput, sizeof(INPUT));
				}
				break;
			}
			case pcKanaLock: {
				BOOL isActive = GetKeyState(VK_KANA) & 0x1;
				if(isActive != static_cast<BOOL>(properties.enabled)) {
					// toggle VK_KANA
					INPUT keyboardInput[2];
					ZeroMemory(keyboardInput, sizeof(keyboardInput));
					keyboardInput[0].type = INPUT_KEYBOARD;
					keyboardInput[0].ki.wVk = VK_KANA;
					keyboardInput[1] = keyboardInput[0];
					keyboardInput[1].ki.dwFlags = KEYEVENTF_KEYUP;
					SendInput(2, keyboardInput, sizeof(INPUT));
				}
				break;
			}
			case pcInsertKey: {
				BYTE keyboardState[256];
				GetKeyboardState(keyboardState);
				if((keyboardState[VK_INSERT] & 0x1) != static_cast<BOOL>(properties.enabled)) {
					// toggle VK_INSERT
					keyboardState[VK_INSERT] = !(GetKeyState(VK_INSERT) & 0x1);
					SetKeyboardState(keyboardState);
				}
				break;
			}
			default:
				m_bRequiresSave = TRUE;
				properties.pOwnerStatBar->SetDirty(TRUE);
		}

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, !simpleMode, 0);
			}

			ApplySettings(TRUE, FALSE);

			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				SendMessage(hWndStatBar, SB_SIMPLE, simpleMode, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_ForeColor(OLE_COLOR newValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
			}

			ApplySettings(TRUE, FALSE);

			if(simpleMode) {
				SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_hIcon(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hIcon);
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_hIcon(OLE_HANDLE newValue)
{
	if(properties.hIcon != LongToHandle(newValue)) {
		properties.hIcon = static_cast<HICON>(LongToHandle(newValue));
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			int panelIndex = properties.panelIndex;
			if(properties.panelIndex == SB_SIMPLEID) {
				panelIndex = -1;
			}

			SendMessage(hWndStatBar, SB_SETICON, panelIndex, reinterpret_cast<LPARAM>(properties.hIcon));
			if(properties.panelIndex == SB_SIMPLEID) {
				InvalidateRect(hWndStatBar, NULL, TRUE);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.panelIndex;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_MinimumWidth(LONG* pValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.minimumWidth;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_MinimumWidth(LONG newValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}
	ATLASSERT(newValue >= 0);
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.minimumWidth != newValue) {
		properties.minimumWidth = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		if(properties.pOwnerStatBar) {
			properties.pOwnerStatBar->ApplyPanelWidths(FALSE);
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_PanelData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.panelData;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_PanelData(LONG newValue)
{
	if(properties.panelData != newValue) {
		properties.panelData = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_ParseTabs(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.parseTabs);
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_ParseTabs(VARIANT_BOOL newValue)
{
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.parseTabs != b) {
		properties.parseTabs = b;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		//if(properties.pOwnerStatBar->IsComctl32Version580OrNewer()) {
			HWND hWndStatBar = properties.GetStatBarHWnd();
			if(IsWindow(hWndStatBar)) {
				BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
				if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
					properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
					SendMessage(hWndStatBar, SB_SIMPLE, !simpleMode, 0);
				}

				ApplySettings(TRUE, FALSE);

				if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
					SendMessage(hWndStatBar, SB_SIMPLE, simpleMode, 0);
					properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
				}
			}
		//}
	}

	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_PreferredWidth(LONG* pValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	ATLASSERT_POINTER(pValue, VARIANT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.preferredWidth;
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_PreferredWidth(LONG newValue)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}
	ATLASSERT(newValue >= 0);
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.preferredWidth != newValue) {
		properties.preferredWidth = newValue;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		if(properties.pOwnerStatBar) {
			properties.pOwnerStatBar->ApplyPanelWidths(FALSE);
		}
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_RightToLeftText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.rightToLeftText);
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_RightToLeftText(VARIANT_BOOL newValue)
{
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.rightToLeftText != b) {
		properties.rightToLeftText = b;
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, !simpleMode, 0);
			}

			ApplySettings(TRUE, FALSE);

			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				SendMessage(hWndStatBar, SB_SIMPLE, simpleMode, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}

	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.text.Copy();
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_Text(BSTR newValue)
{
	if(properties.text != newValue) {
		properties.text.AssignBSTR(newValue);
		switch(properties.content) {
			case pcCapsLock:
			case pcNumLock:
			case pcScrollLock:
			case pcKanaLock:
			case pcInsertKey:
			case pcShortTime:
			case pcLongTime:
			case pcShortDate:
			case pcLongDate:
				break;
			default:
				m_bRequiresSave = TRUE;
				properties.pOwnerStatBar->SetDirty(TRUE);
		}

		HWND hWndStatBar = properties.GetStatBarHWnd();
		if(IsWindow(hWndStatBar)) {
			BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
				SendMessage(hWndStatBar, SB_SIMPLE, !simpleMode, 0);
			}

			ApplySettings(TRUE, FALSE);

			if(simpleMode != (properties.panelIndex == SB_SIMPLEID)) {
				SendMessage(hWndStatBar, SB_SIMPLE, simpleMode, 0);
				properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
			}
		}
	}

	return S_OK;
}

STDMETHODIMP StatusBarPanel::get_ToolTipText(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.toolTipText.Copy();
	return S_OK;
}

STDMETHODIMP StatusBarPanel::put_ToolTipText(BSTR newValue)
{
	if(properties.toolTipText != newValue) {
		properties.toolTipText.AssignBSTR(newValue);
		m_bRequiresSave = TRUE;
		properties.pOwnerStatBar->SetDirty(TRUE);

		if(properties.panelIndex != SB_SIMPLEID) {
			HWND hWndStatBar = properties.GetStatBarHWnd();
			if(IsWindow(hWndStatBar)) {
				TCHAR buffer[MAX_PATH + 1];
				LPTSTR pBuffer = NULL;
				if(lstrlen(COLE2CT(newValue)) > 0) {
					pBuffer = buffer;
					lstrcpyn(pBuffer, COLE2CT(newValue), MAX_PATH);
				}
				SendMessage(hWndStatBar, SB_SETTIPTEXT, properties.panelIndex, reinterpret_cast<LPARAM>(pBuffer));
			}
		}
	}

	return S_OK;
}


STDMETHODIMP StatusBarPanel::GetRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(properties.panelIndex == SB_SIMPLEID) {
		// invalid procedure call - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	HWND hWndStatBar = properties.GetStatBarHWnd();
	ATLASSERT(IsWindow(hWndStatBar));

	BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
	if(simpleMode) {
		properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
		SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
	}
	RECT boundingRectangle = {0};
	LRESULT lr = SendMessage(hWndStatBar, SB_GETRECT, properties.panelIndex, reinterpret_cast<LPARAM>(&boundingRectangle));
	if(simpleMode) {
		SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
		properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
	}

	if(lr) {
		if(pXLeft) {
			*pXLeft = boundingRectangle.left;
		}
		if(pYTop) {
			*pYTop = boundingRectangle.top;
		}
		if(pXRight) {
			*pXRight = boundingRectangle.right;
		}
		if(pYBottom) {
			*pYBottom = boundingRectangle.bottom;
		}

		return S_OK;
	}
	return E_FAIL;
}