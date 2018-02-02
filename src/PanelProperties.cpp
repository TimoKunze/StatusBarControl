// PanelProperties.cpp: The control's "Panels" property page

#include "stdafx.h"
#include "PanelProperties.h"


PanelProperties::PanelProperties()
{
	m_dwTitleID = IDS_TITLEPANELPROPERTIES;
	m_dwDocStringID = IDS_DOCSTRINGPANELPROPERTIES;
}


//////////////////////////////////////////////////////////////////////
// implementation of IPropertyPage
STDMETHODIMP PanelProperties::Apply(void)
{
	ApplySettings();
	return S_OK;
}

STDMETHODIMP PanelProperties::Activate(HWND hWndParent, LPCRECT pRect, BOOL modal)
{
	IPropertyPage2Impl<PanelProperties>::Activate(hWndParent, pRect, modal);

	// attach to the controls
	controls.panelIndexEdit = GetDlgItem(IDC_PANELINDEXEDIT);
	controls.panelIndexUpDown = GetDlgItem(IDC_PANELINDEXUPDOWN);
	controls.insertPanelButton = GetDlgItem(IDC_INSERTPANELBUTTON);
	controls.removePanelButton = GetDlgItem(IDC_REMOVEPANELBUTTON);
	controls.panelTextEdit = GetDlgItem(IDC_PANELTEXTEDIT);
	controls.panelToolTipTextEdit = GetDlgItem(IDC_PANELTOOLTIPTEXTEDIT);
	controls.panelBorderStyleCombo = GetDlgItem(IDC_PANELBORDERSTYLECOMBO);
	controls.panelMinWidthEdit = GetDlgItem(IDC_PANELMINWIDTHEDIT);
	controls.panelAlignmentCombo = GetDlgItem(IDC_PANELALIGNMENTCOMBO);
	controls.panelPrefWidthEdit = GetDlgItem(IDC_PANELPREFWIDTHEDIT);
	controls.panelContentCombo = GetDlgItem(IDC_PANELCONTENTCOMBO);
	controls.panelDataEdit = GetDlgItem(IDC_PANELDATAEDIT);
	controls.panelForeClrRadioDefault = GetDlgItem(IDC_PANELFORECLRRADIODEF);
	controls.panelForeClrRadioSysClr = GetDlgItem(IDC_PANELFORECLRRADIOSYS);
	controls.panelForeClrRadioRGBClr = GetDlgItem(IDC_PANELFORECLRRADIORGB);
	controls.panelForeColorCombo = GetDlgItem(IDC_PANELFORECOLORCOMBO);
	controls.panelForeColorCombo.SetDroppedWidth(165);
	controls.panelForeColorButton.SubclassWindow(GetDlgItem(IDC_PANELFORECOLORBUTTON));
	controls.panelAutoSizeCheck = GetDlgItem(IDC_PANELAUTOSIZECHECK);
	controls.panelEnabledCheck = GetDlgItem(IDC_PANELENABLEDCHECK);
	controls.panelParseTabsCheck = GetDlgItem(IDC_PANELPARSETABSCHECK);
	controls.panelRTLTextCheck = GetDlgItem(IDC_PANELRTLTEXTCHECK);

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

STDMETHODIMP PanelProperties::SetObjects(ULONG objects, IUnknown** ppControls)
{
	if(m_bDirty) {
		Apply();
	}
	IPropertyPage2Impl<PanelProperties>::SetObjects(objects, ppControls);
	LoadSettings();
	return S_OK;
}
// implementation of IPropertyPage
//////////////////////////////////////////////////////////////////////


LRESULT PanelProperties::OnDrawItem(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(wParam != IDC_PANELFORECOLORCOMBO) {
		wasHandled = FALSE;
		return FALSE;
	}

	// TODO: Shouldn't we use the theming API?
	LPDRAWITEMSTRUCT pDrawItemData = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);
	ATLASSERT_POINTER(pDrawItemData, DRAWITEMSTRUCT);
	if(!pDrawItemData) {
		wasHandled = FALSE;
		return FALSE;
	}

	CDCHandle targetDC = pDrawItemData->hDC;
	RECT textRectangle = pDrawItemData->rcItem;
	textRectangle.left = pDrawItemData->rcItem.left + 30;
	if((pDrawItemData->itemID >= 0) && (pDrawItemData->itemAction & (ODA_DRAWENTIRE | ODA_SELECT))) {
		// we will draw the entire item and its selection state has changed
		BOOL enabled = controls.panelForeColorCombo.IsWindowEnabled();

		// set all needed colors
		COLORREF newTextColor = enabled ? GetSysColor(COLOR_WINDOWTEXT) : GetSysColor(COLOR_GRAYTEXT);
		COLORREF oldTextColor = targetDC.SetTextColor(newTextColor);
		COLORREF newBkColor = enabled ? GetSysColor(COLOR_WINDOW) : GetSysColor(COLOR_3DFACE);
		COLORREF oldBkColor = targetDC.SetBkColor(newBkColor);

		if(newTextColor == newBkColor) {
			newTextColor = RGB(0xC0, 0xC0, 0xC0);     // dark gray
		}

		if(enabled) {
			if((pDrawItemData->itemState & ODS_SELECTED) == ODS_SELECTED)	{
				// the item is selected, so use the standard colors
				targetDC.SetTextColor(GetSysColor(COLOR_HIGHLIGHTTEXT));
				targetDC.SetBkColor(GetSysColor(COLOR_HIGHLIGHT));
				// draw the selection rectangle
				targetDC.FillRect(&pDrawItemData->rcItem, GetSysColorBrush(COLOR_HIGHLIGHT));
			} else {
				// clear background
				targetDC.FillRect(&pDrawItemData->rcItem, GetSysColorBrush(COLOR_WINDOW));
			}
		} else {
			// clear background
			targetDC.FillRect(&pDrawItemData->rcItem, GetSysColorBrush(COLOR_3DFACE));
		}

		// now calculate the dimensions of the color rectangle
		RECT colorRectangle = {0};
		colorRectangle.left = pDrawItemData->rcItem.left + 3;
		colorRectangle.top = pDrawItemData->rcItem.top + 2;
		colorRectangle.bottom = pDrawItemData->rcItem.bottom - 2;
		colorRectangle.right = colorRectangle.left + 24;

		// draw the color rectangle
		HBRUSH hPreviousBrush = targetDC.SelectBrush(GetSysColorBrush(pDrawItemData->itemData & 0x000000FF));
		HPEN hBorderPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
		HPEN hPreviousPen = targetDC.SelectPen(hBorderPen);
		targetDC.Rectangle(&colorRectangle);
		targetDC.SelectPen(hPreviousPen);
		targetDC.SelectBrush(hPreviousBrush);

		// now draw the text
		CAtlString itemText;
		controls.panelForeColorCombo.GetLBText(pDrawItemData->itemID, itemText);
		targetDC.DrawText(itemText, itemText.GetLength(), &textRectangle, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_EDITCONTROL);

		targetDC.SetTextColor(oldTextColor);
		targetDC.SetBkColor(oldBkColor);
	}

	if((pDrawItemData->itemAction & ODA_FOCUS) == ODA_FOCUS) {
		targetDC.DrawFocusRect(&pDrawItemData->rcItem);
	}

	return TRUE;
}


LRESULT PanelProperties::OnColorButtonSelChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMCOLORBUTTON pDetails = reinterpret_cast<LPNMCOLORBUTTON>(pNotificationDetails);
	if(pDetails->hdr.idFrom == IDC_PANELFORECOLORBUTTON) {
		flags.changedPanelForeColor = TRUE;
		SetDirty(TRUE);
	}

	return 0;
}

LRESULT PanelProperties::OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT PanelProperties::OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
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

LRESULT PanelProperties::OnUpDownDeltaPosNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMUPDOWN pDetails = reinterpret_cast<LPNMUPDOWN>(pNotificationDetails);
	if(IsPageDirty() == S_OK) {
		ApplyPanelSettings(flags.currentPanelIndex);
	}
	int newPanelIndex = pDetails->iPos + pDetails->iDelta;
	int lowestIndex;
	int highestIndex;
	controls.panelIndexUpDown.GetRange32(lowestIndex, highestIndex);
	if((newPanelIndex >= lowestIndex) && (newPanelIndex <= highestIndex)) {
		LoadPanelSettings(newPanelIndex);
	}

	return 0;
}

LRESULT PanelProperties::OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(TRUE, NULL, NULL, OFN_ENABLESIZING | OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->LoadSettingsFromFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be loaded."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}

LRESULT PanelProperties::OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	ATLASSERT(m_nObjects == 1);

	CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[0]);
	if(pControl) {
		CFileDialog dlg(FALSE, NULL, TEXT("StatusBar Settings.dat"), OFN_ENABLESIZING | OFN_EXPLORER | OFN_CREATEPROMPT | OFN_HIDEREADONLY | OFN_LONGNAMES | OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT, TEXT("All files\0*.*\0\0"), *this);
		if(dlg.DoModal() == IDOK) {
			CComBSTR file = dlg.m_szFileName;

			VARIANT_BOOL b = VARIANT_FALSE;
			pControl->SaveSettingsToFile(file, &b);
			if(b == VARIANT_FALSE) {
				MessageBox(TEXT("The specified file could not be written."), TEXT("Error!"), MB_ICONERROR | MB_OK);
			}
		}
	}
	return 0;
}

LRESULT PanelProperties::OnButtonClicked(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& wasHandled)
{
	switch(controlID) {
		case IDC_INSERTPANELBUTTON: {
			// save the currently selected panel's settings
			if(IsPageDirty() == S_OK) {
				ApplyPanelSettings(flags.currentPanelIndex);
			}

			// insert a panel after the currently selected panel
			for(UINT object = 0; object < m_nObjects; ++object) {
				CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[object]);
				if(pControl) {
					CComPtr<IStatusBarPanels> pPanels = NULL;
					pControl->get_Panels(&pPanels);
					if(pPanels) {
						CComPtr<IStatusBarPanel> pPanel = NULL;
						pPanels->Add(L"", 100, pcText, alLeft, pbsSunken, 0, flags.currentPanelIndex + 1, &pPanel);
					}
				}
			}

			// select the inserted panel
			int lowestIndex;
			int highestIndex;
			controls.panelIndexUpDown.GetRange32(lowestIndex, highestIndex);
			++highestIndex;
			controls.panelIndexUpDown.SetRange32(lowestIndex, highestIndex);
			controls.panelIndexUpDown.SetPos32(flags.currentPanelIndex + 1);
			LoadPanelSettings(controls.panelIndexUpDown.GetPos32());
			break;
		}
		case IDC_REMOVEPANELBUTTON: {
			// remove the currently selected panel
			for(UINT object = 0; object < m_nObjects; ++object) {
				CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[object]);
				if(pControl) {
					CComPtr<IStatusBarPanels> pPanels = NULL;
					pControl->get_Panels(&pPanels);
					if(pPanels) {
						pPanels->Remove(flags.currentPanelIndex);
					}
				}
			}

			// select the next panel
			int lowestIndex;
			int highestIndex;
			controls.panelIndexUpDown.GetRange32(lowestIndex, highestIndex);
			--highestIndex;
			controls.panelIndexUpDown.SetRange32(lowestIndex, highestIndex);
			if(flags.currentPanelIndex > highestIndex) {
				controls.panelIndexUpDown.SetPos32(highestIndex);
			}
			LoadPanelSettings(controls.panelIndexUpDown.GetPos32());
			break;
		}
		case IDC_PANELFORECLRRADIODEF:
		case IDC_PANELFORECLRRADIOSYS:
		case IDC_PANELFORECLRRADIORGB:
			flags.changedPanelForeColor = TRUE;
			SetDirty(TRUE);
			controls.panelForeColorCombo.EnableWindow(controls.panelForeClrRadioSysClr.GetCheck() == BST_CHECKED);
			controls.panelForeColorButton.EnableWindow(controls.panelForeClrRadioRGBClr.GetCheck() == BST_CHECKED);
			break;
		case IDC_PANELAUTOSIZECHECK:
			flags.changedPanelToAutoSize = TRUE;
			SetDirty(TRUE);
			break;
		case IDC_PANELENABLEDCHECK:
			flags.changedPanelEnabled = TRUE;
			SetDirty(TRUE);
			break;
		case IDC_PANELPARSETABSCHECK:
			flags.changedPanelParseTabs = TRUE;
			SetDirty(TRUE);
			break;
		case IDC_PANELRTLTEXTCHECK:
			flags.changedPanelRTLText = TRUE;
			SetDirty(TRUE);
			break;
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT PanelProperties::OnComboSelChange(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	switch(controlID) {
		case IDC_PANELBORDERSTYLECOMBO:
			flags.changedPanelBorderStyle = TRUE;
			break;
		case IDC_PANELALIGNMENTCOMBO:
			flags.changedPanelAlignment = TRUE;
			break;
		case IDC_PANELCONTENTCOMBO:
			flags.changedPanelContent = TRUE;
			controls.panelTextEdit.EnableWindow(controls.panelContentCombo.GetCurSel() == 0);
			break;
		case IDC_PANELFORECOLORCOMBO:
			flags.changedPanelForeColor = TRUE;
			break;
	}
	SetDirty(TRUE);
	return 0;
}

LRESULT PanelProperties::OnEditChange(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	if(controlID != IDC_PANELINDEXEDIT) {
		switch(controlID) {
			case IDC_PANELTEXTEDIT:
				flags.changedPanelText = TRUE;
				break;
			case IDC_PANELTOOLTIPTEXTEDIT:
				flags.changedPanelToolTipText = TRUE;
				break;
			case IDC_PANELMINWIDTHEDIT:
				flags.changedPanelMinimumWidth = TRUE;
				break;
			case IDC_PANELPREFWIDTHEDIT:
				flags.changedPanelPreferredWidth = TRUE;
				break;
			case IDC_PANELDATAEDIT:
				flags.changedPanelData = TRUE;
				break;
		}
		SetDirty(TRUE);
	}
	return 0;
}

LRESULT PanelProperties::OnEditKillFocus(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	if(controlID == IDC_PANELINDEXEDIT) {
		if(IsPageDirty() == S_OK) {
			ApplyPanelSettings(flags.currentPanelIndex);
		}
		LoadPanelSettings(controls.panelIndexUpDown.GetPos32());
	}
	return 0;
}


COLORREF PanelProperties::OLECOLOR2COLORREF(OLE_COLOR oleColor, HPALETTE hPalette/* = NULL*/)
{
	COLORREF color = RGB(0x00, 0x00, 0x00);
	OleTranslateColor(oleColor, hPalette, &color);
	return color;
}

void PanelProperties::ApplySettings(void)
{
	ApplyPanelSettings(flags.currentPanelIndex);
	LoadPanelSettings(flags.currentPanelIndex);
}

void PanelProperties::ApplyPanelSettings(int panelIndex)
{
	// get the controls' contents
	CComBSTR text;
	controls.panelTextEdit.GetWindowText(&text);

	CComBSTR toolTipText;
	controls.panelToolTipTextEdit.GetWindowText(&toolTipText);

	TCHAR buffer[200];
	controls.panelMinWidthEdit.GetWindowText(buffer, 200);
	int minimumWidth = 0;
	ConvertStringToInteger(buffer, minimumWidth);

	controls.panelPrefWidthEdit.GetWindowText(buffer, 200);
	int preferredWidth = 100;
	ConvertStringToInteger(buffer, preferredWidth);

	controls.panelDataEdit.GetWindowText(buffer, 200);
	int panelData = 0;
	ConvertStringToInteger(buffer, panelData);

	PanelBorderStyleConstants borderStyle = static_cast<PanelBorderStyleConstants>(controls.panelBorderStyleCombo.GetCurSel());
	AlignmentConstants alignment = static_cast<AlignmentConstants>(controls.panelAlignmentCombo.GetCurSel());
	PanelContentConstants content = static_cast<PanelContentConstants>(controls.panelContentCombo.GetCurSel());
	if(content == static_cast<PanelContentConstants>(10)) {
		content = pcOwnerDrawn;
	}

	OLE_COLOR foreColor = static_cast<OLE_COLOR>(-2);
	if(controls.panelForeClrRadioDefault.GetCheck() == BST_CHECKED) {
		foreColor = static_cast<OLE_COLOR>(-1);
	} else if(controls.panelForeClrRadioSysClr.GetCheck() == BST_CHECKED) {
		if(controls.panelForeColorCombo.GetCurSel() != -1) {
			foreColor = static_cast<OLE_COLOR>(controls.panelForeColorCombo.GetItemData(controls.panelForeColorCombo.GetCurSel()));
		}
	} else if(controls.panelForeClrRadioRGBClr.GetCheck() == BST_CHECKED) {
		foreColor = controls.panelForeColorButton.GetSelectedColor();
	}

	BOOL autoSized = (controls.panelAutoSizeCheck.GetCheck() == BST_CHECKED);
	BOOL enabled = (controls.panelEnabledCheck.GetCheck() == BST_CHECKED);
	BOOL parseTabs = (controls.panelParseTabsCheck.GetCheck() == BST_CHECKED);
	BOOL rtlText = (controls.panelRTLTextCheck.GetCheck() == BST_CHECKED);

	// apply to all objects
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[object]);
		if(pControl) {
			CComPtr<IStatusBarPanel> pPanel = NULL;

			if(panelIndex == -1) {
				// the simple mode panel
				pControl->get_SimplePanel(&pPanel);
			} else {
				// a "real" panel
				CComPtr<IStatusBarPanels> pPanels = NULL;
				pControl->get_Panels(&pPanels);
				if(pPanels) {
					pPanels->get_Item(panelIndex, &pPanel);
				}
			}

			if(pPanel) {
				if(flags.changedPanelText) {
					pPanel->put_Text(text.Detach());
				}
				if(flags.changedPanelToolTipText) {
					pPanel->put_ToolTipText(toolTipText.Detach());
				}
				if(flags.changedPanelBorderStyle) {
					if(borderStyle != static_cast<PanelBorderStyleConstants>(-1)) {
						pPanel->put_BorderStyle(borderStyle);
					}
				}
				if(flags.changedPanelData) {
					pPanel->put_PanelData(panelData);
				}
				if(flags.changedPanelParseTabs) {
					pPanel->put_ParseTabs(BOOL2VARIANTBOOL(parseTabs));
				}
				if(flags.changedPanelRTLText) {
					pPanel->put_RightToLeftText(BOOL2VARIANTBOOL(rtlText));
				}

				if(panelIndex != -1) {
					if(flags.changedPanelAlignment) {
						if(alignment != static_cast<AlignmentConstants>(-1)) {
							pPanel->put_Alignment(alignment);
						}
					}
					if(flags.changedPanelContent) {
						if(content != static_cast<PanelContentConstants>(-1)) {
							pPanel->put_Content(content);
						}
					}
					if(flags.changedPanelMinimumWidth) {
						pPanel->put_MinimumWidth(minimumWidth);
					}
					if(flags.changedPanelPreferredWidth) {
						pPanel->put_PreferredWidth(preferredWidth);
					}
					if(flags.changedPanelForeColor) {
						pPanel->put_ForeColor(foreColor);
					}
					if(flags.changedPanelToAutoSize) {
						if(autoSized) {
							pControl->putref_PanelToAutoSize(pPanel);
						} else {
							pControl->putref_PanelToAutoSize(NULL);
						}
					}
					if(flags.changedPanelEnabled) {
						pPanel->put_Enabled(BOOL2VARIANTBOOL(enabled));
					}
				}
			} else {
				// panel not found - abort
				return;
			}
		}
	}

	flags.changedPanelText = FALSE;
	flags.changedPanelToolTipText = FALSE;
	flags.changedPanelBorderStyle = FALSE;
	flags.changedPanelMinimumWidth = FALSE;
	flags.changedPanelAlignment = FALSE;
	flags.changedPanelPreferredWidth = FALSE;
	flags.changedPanelContent = FALSE;
	flags.changedPanelData = FALSE;
	flags.changedPanelForeColor = FALSE;
	flags.changedPanelToAutoSize = FALSE;
	flags.changedPanelEnabled = FALSE;
	flags.changedPanelParseTabs = FALSE;
	flags.changedPanelRTLText = FALSE;
	SetDirty(FALSE);
}

void PanelProperties::LoadSettings(void)
{
	if(!controls.toolbar.IsWindow()) {
		// this will happen in Visual Studio's dialog editor if settings are loaded from a file
		return;
	}

	controls.toolbar.EnableButton(ID_LOADSETTINGS, (m_nObjects == 1));
	controls.toolbar.EnableButton(ID_SAVESETTINGS, (m_nObjects == 1));

	// get the settings
	int minPanelCount = INT_MAX;
	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[object]);
		if(pControl) {
			CComPtr<IStatusBarPanels> pPanels = NULL;
			pControl->get_Panels(&pPanels);
			if(pPanels) {
				LONG l = 0;
				pPanels->Count(&l);
				minPanelCount = min(minPanelCount, l);
			}
		}
	}

	// fill the controls
	controls.panelIndexUpDown.SetRange32(-1, minPanelCount - 1);

	controls.panelBorderStyleCombo.ResetContent();
	controls.panelBorderStyleCombo.AddString(TEXT("0 - Sunken"));
	controls.panelBorderStyleCombo.AddString(TEXT("1 - None"));
	controls.panelBorderStyleCombo.AddString(TEXT("2 - Raised"));

	controls.panelAlignmentCombo.ResetContent();
	controls.panelAlignmentCombo.AddString(TEXT("0 - Left"));
	controls.panelAlignmentCombo.AddString(TEXT("1 - Center"));
	controls.panelAlignmentCombo.AddString(TEXT("2 - Right"));

	controls.panelContentCombo.ResetContent();
	controls.panelContentCombo.AddString(TEXT("0 - Text"));
	controls.panelContentCombo.AddString(TEXT("1 - Caps Lock"));
	controls.panelContentCombo.AddString(TEXT("2 - Num Lock"));
	controls.panelContentCombo.AddString(TEXT("3 - Scroll Lock"));
	controls.panelContentCombo.AddString(TEXT("4 - Kana Lock"));
	controls.panelContentCombo.AddString(TEXT("5 - Insert Key"));
	controls.panelContentCombo.AddString(TEXT("6 - Short Time"));
	controls.panelContentCombo.AddString(TEXT("7 - Long Time"));
	controls.panelContentCombo.AddString(TEXT("8 - Short Date"));
	controls.panelContentCombo.AddString(TEXT("9 - Long Date"));
	controls.panelContentCombo.AddString(TEXT("99 - Ownerdrawn"));

	controls.panelForeColorCombo.ResetContent();
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ActiveBorder")), 0x8000000A);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ActiveCaption")), 0x80000002);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ActiveCaptionText")), 0x80000009);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ActiveCaptionGradient")), 0x8000001B);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("AppWorkSpace")), 0x8000000C);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Control")), 0x8000000F);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ControlDark")), 0x80000010);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ControlDarkDark")), 0x80000015);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ControlLight")), 0x80000016);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ControlLightLight")), 0x80000014);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ControlText")), 0x80000012);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Desktop")), 0x80000001);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("GrayText")), 0x80000011);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Highlight")), 0x8000000D);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("HighlightText")), 0x8000000E);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Hot")), 0x8000001A);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("InactiveBorder")), 0x8000000B);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("InactiveCaption")), 0x80000003);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("InactiveCaptionGradient")), 0x8000001C);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("InactiveCaptionText")), 0x80000013);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Info")), 0x80000018);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("InfoText")), 0x80000017);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Menu")), 0x80000004);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("MenuBar")), 0x8000001E);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("MenuHighlight")), 0x8000001D);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("MenuText")), 0x80000007);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("ScrollBar")), 0x80000000);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("Window")), 0x80000005);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("WindowFrame")), 0x80000006);
	controls.panelForeColorCombo.SetItemData(controls.panelForeColorCombo.AddString(TEXT("WindowText")), 0x80000008);

	controls.panelForeColorButton.SetDefaultColor(RGB(0, 0, 0));
	controls.panelForeColorButton.SetDefaultColorText(TEXT(""));

	// select the simple mode panel
	controls.panelIndexUpDown.SetPos32(flags.currentPanelIndex);
	LoadPanelSettings(flags.currentPanelIndex);

	SetDirty(FALSE);
}

void PanelProperties::LoadPanelSettings(int panelIndex)
{
	controls.removePanelButton.EnableWindow(panelIndex != -1);
	controls.panelAlignmentCombo.EnableWindow(panelIndex != -1);
	controls.panelMinWidthEdit.EnableWindow(panelIndex != -1);
	controls.panelPrefWidthEdit.EnableWindow(panelIndex != -1);
	controls.panelContentCombo.EnableWindow(panelIndex != -1);
	controls.panelForeClrRadioDefault.EnableWindow(panelIndex != -1);
	controls.panelForeClrRadioSysClr.EnableWindow(panelIndex != -1);
	controls.panelForeClrRadioRGBClr.EnableWindow(panelIndex != -1);
	controls.panelForeColorCombo.EnableWindow(panelIndex != -1);
	controls.panelForeColorButton.EnableWindow(panelIndex != -1);
	controls.panelAutoSizeCheck.EnableWindow(panelIndex != -1);
	controls.panelEnabledCheck.EnableWindow(panelIndex != -1);

	CComBSTR text;
	CComBSTR toolTipText;
	PanelBorderStyleConstants borderStyle = static_cast<PanelBorderStyleConstants>(-1);
	AlignmentConstants alignment = static_cast<AlignmentConstants>(-1);
	PanelContentConstants content = static_cast<PanelContentConstants>(-1);
	BOOL minWidthEmpty = FALSE;
	int minimumWidth = 0;
	BOOL prefWidthEmpty = FALSE;
	int preferredWidth = -1;
	BOOL panelDataEmpty = FALSE;
	int panelData = 0;
	OLE_COLOR foreColor = static_cast<OLE_COLOR>(-2);
	int autoSized = BST_INDETERMINATE;
	int enabled = BST_INDETERMINATE;
	int parseTabs = BST_INDETERMINATE;
	int rtlText = BST_INDETERMINATE;

	for(UINT object = 0; object < m_nObjects; ++object) {
		CComQIPtr<IStatusBar, &IID_IStatusBar> pControl(m_ppUnk[object]);
		if(pControl) {
			CComPtr<IStatusBarPanel> pPanel = NULL;

			if(panelIndex == -1) {
				// the simple mode panel
				pControl->get_SimplePanel(&pPanel);
			} else {
				// a "real" panel
				CComPtr<IStatusBarPanels> pPanels = NULL;
				pControl->get_Panels(&pPanels);
				if(pPanels) {
					pPanels->get_Item(panelIndex, &pPanel);
				}
			}

			if(pPanel) {
				// retrieve the settings
				BSTR bstr;
				pPanel->get_Text(&bstr);
				if(object == 0) {
					text = bstr;
				} else {
					if(lstrcmp(COLE2CT(text), COLE2CT(bstr)) != 0) {
						text.Empty();
					}
				}
				SysFreeString(bstr);

				pPanel->get_ToolTipText(&bstr);
				if(object == 0) {
					toolTipText = bstr;
				} else {
					if(lstrcmp(COLE2CT(toolTipText), COLE2CT(bstr)) != 0) {
						toolTipText.Empty();
					}
				}

				PanelBorderStyleConstants pbs;
				pPanel->get_BorderStyle(&pbs);
				if(object == 0) {
					borderStyle = pbs;
				} else {
					if(borderStyle != pbs) {
						borderStyle = static_cast<PanelBorderStyleConstants>(-1);
					}
				}

				LONG l;
				pPanel->get_PanelData(&l);
				if(object == 0) {
					panelData = l;
				} else {
					if(panelData != l) {
						panelDataEmpty = TRUE;
					}
				}

				VARIANT_BOOL b;
				pPanel->get_ParseTabs(&b);
				if(object == 0) {
					parseTabs = (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED);
				} else {
					if(parseTabs != (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED)) {
						parseTabs = BST_INDETERMINATE;
					}
				}

				pPanel->get_RightToLeftText(&b);
				if(object == 0) {
					rtlText = (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED);
				} else {
					if(rtlText != (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED)) {
						rtlText = BST_INDETERMINATE;
					}
				}

				if(panelIndex == -1) {
					alignment = alLeft;
					content = pcText;
					minimumWidth = 0;
					preferredWidth = -1;
					foreColor = static_cast<OLE_COLOR>(-1);
					autoSized = BST_CHECKED;
					enabled = BST_CHECKED;
				} else {
					AlignmentConstants al;
					pPanel->get_Alignment(&al);
					if(object == 0) {
						alignment = al;
					} else {
						if(alignment != al) {
							alignment = static_cast<AlignmentConstants>(-1);
						}
					}

					PanelContentConstants pc;
					pPanel->get_Content(&pc);
					if(object == 0) {
						content = pc;
					} else {
						if(content != pc) {
							content = static_cast<PanelContentConstants>(-1);
						}
					}

					pPanel->get_MinimumWidth(&l);
					if(object == 0) {
						minimumWidth = l;
					} else {
						if(minimumWidth != l) {
							minWidthEmpty = TRUE;
						}
					}

					pPanel->get_PreferredWidth(&l);
					if(object == 0) {
						preferredWidth = l;
					} else {
						if(preferredWidth != l) {
							prefWidthEmpty = TRUE;
						}
					}

					OLE_COLOR clr;
					pPanel->get_ForeColor(&clr);
					if(object == 0) {
						foreColor = clr;
					} else {
						if(foreColor != clr) {
							foreColor = static_cast<OLE_COLOR>(-2);
						}
					}

					CComPtr<IStatusBarPanel> pAutoSizedPanel = NULL;
					pControl->get_PanelToAutoSize(&pAutoSizedPanel);
					if(pAutoSizedPanel) {
						pAutoSizedPanel->get_Index(&l);
						b = BOOL2VARIANTBOOL(l == panelIndex);
					} else {
						b = VARIANT_FALSE;
					}
					if(object == 0) {
						autoSized = (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED);
					} else {
						if(autoSized != (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED)) {
							autoSized = BST_INDETERMINATE;
						}
					}

					pPanel->get_Enabled(&b);
					if(object == 0) {
						enabled = (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED);
					} else {
						if(enabled != (VARIANTBOOL2BOOL(b) ? BST_CHECKED : BST_UNCHECKED)) {
							enabled = BST_INDETERMINATE;
						}
					}
				}
			} else {
				// panel not found - abort
				return;
			}
		}
	}

	// fill the controls
	LPTSTR pBuffer = ConvertIntegerToString(panelIndex);
	controls.panelIndexEdit.SetWindowText(pBuffer);
	SECUREFREE(pBuffer);
	controls.panelTextEdit.SetWindowText(COLE2CT(text));
	controls.panelToolTipTextEdit.SetWindowText(COLE2CT(toolTipText));
	controls.panelBorderStyleCombo.SetCurSel(borderStyle);
	controls.panelAlignmentCombo.SetCurSel(alignment);
	controls.panelContentCombo.SetCurSel((content == pcOwnerDrawn ? 10 : content));
	controls.panelTextEdit.EnableWindow(controls.panelContentCombo.GetCurSel() == 0);
	if(minWidthEmpty) {
		controls.panelMinWidthEdit.SetWindowText(NULL);
	} else {
		pBuffer = ConvertIntegerToString(minimumWidth);
		controls.panelMinWidthEdit.SetWindowText(pBuffer);
		SECUREFREE(pBuffer);
	}
	if(prefWidthEmpty) {
		controls.panelPrefWidthEdit.SetWindowText(NULL);
	} else {
		pBuffer = ConvertIntegerToString(preferredWidth);
		controls.panelPrefWidthEdit.SetWindowText(pBuffer);
		SECUREFREE(pBuffer);
	}
	if(panelDataEmpty) {
		controls.panelDataEdit.SetWindowText(NULL);
	} else {
		pBuffer = ConvertIntegerToString(panelData);
		controls.panelDataEdit.SetWindowText(pBuffer);
		SECUREFREE(pBuffer);
	}

	controls.panelForeClrRadioDefault.SetCheck(BST_UNCHECKED);
	controls.panelForeClrRadioSysClr.SetCheck(BST_UNCHECKED);
	controls.panelForeClrRadioRGBClr.SetCheck(BST_UNCHECKED);
	if(foreColor != -2) {
		if(foreColor == -1) {
			controls.panelForeClrRadioDefault.SetCheck(BST_CHECKED);
			for(int i = 0; i < controls.panelForeColorCombo.GetCount(); ++i) {
				if(controls.panelForeColorCombo.GetItemData(i) == 0x80000012) {
					controls.panelForeColorCombo.SetCurSel(i);
					break;
				}
			}
			controls.panelForeColorButton.SetSelectedColor(RGB(0, 0, 0));
		} else if(foreColor & 0xFF000000) {
			controls.panelForeClrRadioSysClr.SetCheck(BST_CHECKED);
			for(int i = 0; i < controls.panelForeColorCombo.GetCount(); ++i) {
				if(controls.panelForeColorCombo.GetItemData(i) == foreColor) {
					controls.panelForeColorCombo.SetCurSel(i);
					break;
				}
			}
			controls.panelForeColorButton.SetSelectedColor(OLECOLOR2COLORREF(foreColor));
		} else {
			controls.panelForeClrRadioRGBClr.SetCheck(BST_CHECKED);
			controls.panelForeColorButton.SetSelectedColor(OLECOLOR2COLORREF(foreColor));
			for(int i = 0; i < controls.panelForeColorCombo.GetCount(); ++i) {
				if(controls.panelForeColorCombo.GetItemData(i) == 0x80000012) {
					controls.panelForeColorCombo.SetCurSel(i);
					break;
				}
			}
		}
	}
	controls.panelForeColorCombo.EnableWindow(controls.panelForeClrRadioSysClr.GetCheck() == BST_CHECKED);
	controls.panelForeColorButton.EnableWindow(controls.panelForeClrRadioRGBClr.GetCheck() == BST_CHECKED);

	controls.panelAutoSizeCheck.SetCheck(autoSized);
	controls.panelEnabledCheck.SetCheck(enabled);
	controls.panelParseTabsCheck.SetCheck(parseTabs);
	controls.panelRTLTextCheck.SetCheck(rtlText);

	flags.changedPanelText = FALSE;
	flags.changedPanelToolTipText = FALSE;
	flags.changedPanelBorderStyle = FALSE;
	flags.changedPanelMinimumWidth = FALSE;
	flags.changedPanelAlignment = FALSE;
	flags.changedPanelPreferredWidth = FALSE;
	flags.changedPanelContent = FALSE;
	flags.changedPanelData = FALSE;
	flags.changedPanelForeColor = FALSE;
	flags.changedPanelToAutoSize = FALSE;
	flags.changedPanelEnabled = FALSE;
	flags.changedPanelParseTabs = FALSE;
	flags.changedPanelRTLText = FALSE;
	flags.currentPanelIndex = panelIndex;
	SetDirty(FALSE);
}