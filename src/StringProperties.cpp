// StringProperties.cpp: The control's "String Properties" property page

#include "stdafx.h"
#include "StringProperties.h"


StringProperties::StringProperties()
{
	m_dwTitleID = IDS_TITLESTRINGPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGSTRINGPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP StringProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP StringProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<StringProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.propertyCombo.Attach(GetDlgItem(IDC_PROPERTYCOMBO));
	controls.propertyEdit.Attach(GetDlgItem(IDC_PROPERTYEDIT));

	// setup the toolbar
	CRect toolbarRect;
	GetClientRect(&toolbarRect);
	toolbarRect.OffsetRect(0, 2);
	toolbarRect.left += toolbarRect.right - 46;
	toolbarRect.bottom = toolbarRect.top + 22;
	controls.toolbar.Create(*this, toolbarRect, NULL, WS_CHILDWINDOW | WS_VISIBLE | TBSTYLE_TRANSPARENT | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE, 0);
	controls.toolbar.SetButtonStructSize();
	controls.imagelistEnabled.CreateFromImage(IDB_TOOLBARENABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetImageList(controls.imagelistEnabled);
	controls.imagelistDisabled.CreateFromImage(IDB_TOOLBARDISABLED, 16, 0, RGB(255, 0, 255), IMAGE_BITMAP, LR_CREATEDIBSECTION);
	controls.toolbar.SetDisabledImageList(controls.imagelistDisabled);

	// insert the buttons
	TBBUTTON buttons[2];
	ZeroMemory(buttons, sizeof(buttons));
	buttons[0].iBitmap = 0;
	buttons[0].idCommand = ID_LOADSETTINGS;
	buttons[0].fsState = TBSTATE_ENABLED;
	buttons[0].fsStyle = BTNS_BUTTON;
	buttons[1].iBitmap = 1;
	buttons[1].idCommand = ID_SAVESETTINGS;
	buttons[1].fsStyle = BTNS_BUTTON;
	buttons[1].fsState = TBSTATE_ENABLED;
	controls.toolbar.AddButtons(2, buttons);

	LoadSettings();
	return S_OK;
}

STDMETHODIMP StringProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<StringProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage2
STDMETHODIMP StringProperties::EditProperty(DISPID dispID)
{
	properties.propertyToEdit = dispID;
	switch(properties.selectedPropertyItemID) {
		case 0:
			controls.propertyEdit.GetWindowText(properties.customCapsLockText);
			break;
		case 1:
			controls.propertyEdit.GetWindowText(properties.customInsertKeyText);
			break;
		case 2:
			controls.propertyEdit.GetWindowText(properties.customKanaLockText);
			break;
		case 3:
			controls.propertyEdit.GetWindowText(properties.customNumLockText);
			break;
		case 4:
			controls.propertyEdit.GetWindowText(properties.customScrollLockText);
			break;
	}

	LPUNKNOWN pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IStatusBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		switch(properties.propertyToEdit) {
			case DISPID_STATBAR_CUSTOMCAPSLOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 0) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMINSERTKEYTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 1) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMKANALOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 2) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMNUMLOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 3) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 4) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
		}
		pControl->Release();
	}

	properties.selectedPropertyItemID = -1;
	properties.selectedPropertyItemID = static_cast<int>(controls.propertyCombo.GetItemData(0));
	switch(properties.selectedPropertyItemID) {
		case 0:
			controls.propertyEdit.SetWindowText(properties.customCapsLockText);
			break;
		case 1:
			controls.propertyEdit.SetWindowText(properties.customInsertKeyText);
			break;
		case 2:
			controls.propertyEdit.SetWindowText(properties.customKanaLockText);
			break;
		case 3:
			controls.propertyEdit.SetWindowText(properties.customNumLockText);
			break;
		case 4:
			controls.propertyEdit.SetWindowText(properties.customScrollLockText);
			break;
		default:
			controls.propertyEdit.SetWindowText(TEXT(""));
			break;
	}
	return S_OK;
}
// implementation of IPropertyPage2
//////////////////////////////////////////////////////////////////////


LRESULT StringProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOA pDetails = reinterpret_cast<LPNMTTDISPINFOA>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEA(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT StringProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTTDISPINFOW pDetails = reinterpret_cast<LPNMTTDISPINFOW>(pNotificationDetails);
	pDetails->hinst = ModuleHelper::GetResourceInstance();
	switch(pDetails->hdr.idFrom) {
		case ID_LOADSETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_LOADSETTINGS);
			break;
		case ID_SAVESETTINGS:
			pDetails->lpszText = MAKEINTRESOURCEW(IDS_SAVESETTINGS);
			break;
	}
	return 0;
}

LRESULT StringProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IStatusBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IStatusBar*>(pControl)->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT StringProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	IUnknown* pControl = NULL;
	if(m_ppUnk[0]->QueryInterface(IID_IStatusBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
		CFileDialog dlg(FALSE, NULL, TEXT("StatusBar Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			reinterpret_cast<IStatusBar*>(pControl)->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
		pControl->Release();
	}
	return 0;
}

LRESULT StringProperties::OnPropertySelChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	switch(properties.selectedPropertyItemID) {
		case 0:
			controls.propertyEdit.GetWindowText(properties.customCapsLockText);
			break;
		case 1:
			controls.propertyEdit.GetWindowText(properties.customInsertKeyText);
			break;
		case 2:
			controls.propertyEdit.GetWindowText(properties.customKanaLockText);
			break;
		case 3:
			controls.propertyEdit.GetWindowText(properties.customNumLockText);
			break;
		case 4:
			controls.propertyEdit.GetWindowText(properties.customScrollLockText);
			break;
	}

	HRESULT dirty = IsPageDirty();

	int itemIndex = controls.propertyCombo.GetCurSel();
	properties.selectedPropertyItemID = -1;
	if(itemIndex != CB_ERR) {
		properties.selectedPropertyItemID = static_cast<int>(controls.propertyCombo.GetItemData(itemIndex));
	}
	switch(properties.selectedPropertyItemID) {
		case 0:
			controls.propertyEdit.SetWindowText(properties.customCapsLockText);
			break;
		case 1:
			controls.propertyEdit.SetWindowText(properties.customInsertKeyText);
			break;
		case 2:
			controls.propertyEdit.SetWindowText(properties.customKanaLockText);
			break;
		case 3:
			controls.propertyEdit.SetWindowText(properties.customNumLockText);
			break;
		case 4:
			controls.propertyEdit.SetWindowText(properties.customScrollLockText);
			break;
		default:
			controls.propertyEdit.SetWindowText(TEXT(""));
			break;
	}

	SetDirty(dirty == S_OK);
	return 0;
}

LRESULT StringProperties::OnChange(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	SetDirty(TRUE);
	return 0;
}


void StringProperties::ApplySettings(void)
{
	switch(properties.selectedPropertyItemID) {
		case 0:
			controls.propertyEdit.GetWindowText(properties.customCapsLockText);
			break;
		case 1:
			controls.propertyEdit.GetWindowText(properties.customInsertKeyText);
			break;
		case 2:
			controls.propertyEdit.GetWindowText(properties.customKanaLockText);
			break;
		case 3:
			controls.propertyEdit.GetWindowText(properties.customNumLockText);
			break;
		case 4:
			controls.propertyEdit.GetWindowText(properties.customScrollLockText);
			break;
	}

	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IStatusBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			CComBSTR bstr = properties.customCapsLockText;
			reinterpret_cast<IStatusBar*>(pControl)->put_CustomCapsLockText(bstr);
			bstr = properties.customInsertKeyText;
			reinterpret_cast<IStatusBar*>(pControl)->put_CustomInsertKeyText(bstr);
			bstr = properties.customKanaLockText;
			reinterpret_cast<IStatusBar*>(pControl)->put_CustomKanaLockText(bstr);
			bstr = properties.customNumLockText;
			reinterpret_cast<IStatusBar*>(pControl)->put_CustomNumLockText(bstr);
			bstr = properties.customScrollLockText;
			reinterpret_cast<IStatusBar*>(pControl)->put_CustomScrollLockText(bstr);
			pControl->Release();
		}
	}

	SetDirty(FALSE);
}

void StringProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	properties.customCapsLockText = TEXT("");
	properties.customInsertKeyText = TEXT("");
	properties.customKanaLockText = TEXT("");
	properties.customNumLockText = TEXT("");
	properties.customScrollLockText = TEXT("");
	BSTR value;
	VARIANT_BOOL rtlLayout = VARIANT_FALSE;
	for(UINT object = 0; object < m_nObjects; ++object) {
		LPUNKNOWN pControl = NULL;
		if(m_ppUnk[object]->QueryInterface(IID_IStatusBar, reinterpret_cast<LPVOID*>(&pControl)) == S_OK) {
			if(m_nObjects == 1) {
				reinterpret_cast<IStatusBar*>(pControl)->get_CustomCapsLockText(&value);
				properties.customCapsLockText = value;
				SysFreeString(value);
				reinterpret_cast<IStatusBar*>(pControl)->get_CustomInsertKeyText(&value);
				properties.customInsertKeyText = value;
				SysFreeString(value);
				reinterpret_cast<IStatusBar*>(pControl)->get_CustomKanaLockText(&value);
				properties.customKanaLockText = value;
				SysFreeString(value);
				reinterpret_cast<IStatusBar*>(pControl)->get_CustomNumLockText(&value);
				properties.customNumLockText = value;
				SysFreeString(value);
				reinterpret_cast<IStatusBar*>(pControl)->get_CustomScrollLockText(&value);
				properties.customScrollLockText = value;
				SysFreeString(value);
				reinterpret_cast<IStatusBar*>(pControl)->get_RightToLeftLayout(&rtlLayout);
			}
			pControl->Release();
		}
	}

	// fill the controls
	controls.propertyEdit.SetWindowText(TEXT(""));
	if(rtlLayout != VARIANT_FALSE) {
		controls.propertyEdit.ModifyStyleEx(0, WS_EX_LAYOUTRTL);
	} else {
		controls.propertyEdit.ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
	}
	controls.propertyCombo.ResetContent();
	int i = controls.propertyCombo.AddString(TEXT("CustomCapsLockText"));
	controls.propertyCombo.SetItemData(i, 0);
	i = controls.propertyCombo.AddString(TEXT("CustomInsertKeyText"));
	controls.propertyCombo.SetItemData(i, 1);
	i = controls.propertyCombo.AddString(TEXT("CustomKanaLockText"));
	controls.propertyCombo.SetItemData(i, 2);
	i = controls.propertyCombo.AddString(TEXT("CustomNumLockText"));
	controls.propertyCombo.SetItemData(i, 3);
	i = controls.propertyCombo.AddString(TEXT("CustomScrollLockText"));
	controls.propertyCombo.SetItemData(i, 4);
	if(controls.propertyCombo.GetCount() > 0) {
		switch(properties.propertyToEdit) {
			case DISPID_STATBAR_CUSTOMCAPSLOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 0) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMINSERTKEYTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 1) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMKANALOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 2) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMNUMLOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 3) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			case DISPID_STATBAR_CUSTOMSCROLLLOCKTEXT:
				for(int i = 0; i < controls.propertyCombo.GetCount(); i++) {
					if(static_cast<int>(controls.propertyCombo.GetItemData(i)) == 4) {
						controls.propertyCombo.SetCurSel(i);
						break;
					}
				}
				break;
			default:
				controls.propertyCombo.SetCurSel(0);
				break;
		}
		properties.selectedPropertyItemID = -1;
		properties.selectedPropertyItemID = static_cast<int>(controls.propertyCombo.GetItemData(controls.propertyCombo.GetCurSel()));
		switch(properties.selectedPropertyItemID) {
			case 0:
				controls.propertyEdit.SetWindowText(properties.customCapsLockText);
				break;
			case 1:
				controls.propertyEdit.SetWindowText(properties.customInsertKeyText);
				break;
			case 2:
				controls.propertyEdit.SetWindowText(properties.customKanaLockText);
				break;
			case 3:
				controls.propertyEdit.SetWindowText(properties.customNumLockText);
				break;
			case 4:
				controls.propertyEdit.SetWindowText(properties.customScrollLockText);
				break;
			default:
				controls.propertyEdit.SetWindowText(TEXT(""));
				break;
		}
	}

	SetDirty(FALSE);
}