// ColorButton.cpp: a button control that can be used as color picker

#include "stdafx.h"
#include "helpers.h"
#include "ColorButton.h"


// initialize complex constants
ColorButton::ColorTableEntry ColorButton::predefinedColors[] = {
    {RGB(0x00, 0x00, 0x00), TEXT("Black")},
    {RGB(0xA5, 0x2A, 0x00), TEXT("Brown")},
    {RGB(0x00, 0x40, 0x40), TEXT("Dark Olive Green")},
    {RGB(0x00, 0x55, 0x00), TEXT("Dark Green")},
    {RGB(0x00, 0x00, 0x5E), TEXT("Dark Teal")},
    {RGB(0x00, 0x00, 0x8B), TEXT("Dark Blue")},
    {RGB(0x4B, 0x00, 0x82), TEXT("Indigo")},
    {RGB(0x28, 0x28, 0x28), TEXT("Dark Grey")},

    {RGB(0x8B, 0x00, 0x00), TEXT("Dark Red")},
    {RGB(0xFF, 0x68, 0x20), TEXT("Orange")},
    {RGB(0x8B, 0x8B, 0x00), TEXT("Dark Yellow")},
    {RGB(0x00, 0x93, 0x00), TEXT("Green")},
    {RGB(0x38, 0x8E, 0x8E), TEXT("Teal")},
    {RGB(0x00, 0x00, 0xFF), TEXT("Blue")},
    {RGB(0x7B, 0x7B, 0xC0), TEXT("Blue-Grey")},
    {RGB(0x66, 0x66, 0x66), TEXT("Grey - 40")},

    {RGB(0xFF, 0x00, 0x00), TEXT("Red")},
    {RGB(0xFF, 0xAD, 0x5B), TEXT("Light Orange")},
    {RGB(0x32, 0xCD, 0x32), TEXT("Lime")},
    {RGB(0x3C, 0xB3, 0x71), TEXT("Sea Green")},
    {RGB(0x7F, 0xFF, 0xD4), TEXT("Aqua")},
    {RGB(0x7D, 0x9E, 0xC0), TEXT("Light Blue")},
    {RGB(0x80, 0x00, 0x80), TEXT("Violet")},
    {RGB(0x7F, 0x7F, 0x7F), TEXT("Grey - 50")},

    {RGB(0xFF, 0xC0, 0xCB), TEXT("Pink")},
    {RGB(0xFF, 0xD7, 0x00), TEXT("Gold")},
    {RGB(0xFF, 0xFF, 0x00), TEXT("Yellow")},
    {RGB(0x00, 0xFF, 0x00), TEXT("Bright Green")},
    {RGB(0x40, 0xE0, 0xD0), TEXT("Turquoise")},
    {RGB(0xC0, 0xFF, 0xFF), TEXT("Skyblue")},
    {RGB(0x48, 0x00, 0x48), TEXT("Plum")},
    {RGB(0xC0, 0xC0, 0xC0), TEXT("Light Grey")},

    {RGB(0xFF, 0xE4, 0xE1), TEXT("Rose")},
    {RGB(0xD2, 0xB4, 0x8C), TEXT("Tan")},
    {RGB(0xFF, 0xFF, 0xE0), TEXT("Light Yellow")},
    {RGB(0x98, 0xFB, 0x98), TEXT("Pale Green ")},
    {RGB(0xAF, 0xEE, 0xEE), TEXT("Pale Turquoise")},
    {RGB(0x68, 0x83, 0x8B), TEXT("Pale Blue")},
    {RGB(0xE6, 0xE6, 0xFA), TEXT("Lavender")},
    {RGB(0xFF, 0xFF, 0xFF), TEXT("White")}
};

const WTL::CSize ColorButton::innerTextCellBorderSize(3, 3);
const WTL::CSize ColorButton::outerTextCellBorderSize(2, 2);
const WTL::CSize ColorButton::innerColorCellBorderSize(2, 2);
const WTL::CSize ColorButton::outerColorCellBorderSize(0, 0);
const WTL::CSize ColorButton::colorCellContentSize(14, 14);


#pragma warning(push)
#pragma warning(disable: 4355)     // passing "this" to a member constructor
ColorButton::ColorButton() :
    dropDownWindow(this, 1)
{
}
#pragma warning(pop)


HWND ColorButton::Create(HWND hWndParent, RECT& windowRect, LPCTSTR lpWindowName/* = NULL*/, DWORD windowStyle/* = 0*/, DWORD extWindowStyle/* = 0*/, UINT controlID/* = 0*/, LPVOID lpCreateParam/* = NULL*/)
{
	HWND ret = CWindow::Create(WC_BUTTON, hWndParent, windowRect, lpWindowName, windowStyle | BS_OWNERDRAW, extWindowStyle, controlID, lpCreateParam);

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	if(IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				OpenThemeData(VSCLASS_BUTTON);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return ret;
}

HWND ColorButton::Create(HWND hWndParent, LPRECT lpWindowRect/* = NULL*/, LPCTSTR lpWindowName/* = NULL*/, DWORD windowStyle/* = 0*/, DWORD extWindowStyle/* = 0*/, HMENU hMenu/* = NULL*/, LPVOID lpCreateParam/* = NULL*/)
{
	HWND ret = CWindow::Create(WC_BUTTON, hWndParent, lpWindowRect, lpWindowName, windowStyle | BS_OWNERDRAW, extWindowStyle, hMenu, lpCreateParam);

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	if(IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				OpenThemeData(VSCLASS_BUTTON);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return ret;
}

BOOL ColorButton::SubclassWindow(HWND hWnd)
{
	BOOL ret = CWindowImpl<ColorButton>::SubclassWindow(hWnd);
	ModifyStyle(0, BS_OWNERDRAW);

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	if(IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				OpenThemeData(VSCLASS_BUTTON);
			}
			FreeLibrary(hThemeDLL);
		}
	}

	return ret;
}


COLORREF ColorButton::GetSelectedColor(void) const
{
	return properties.selectedColor;
}

void ColorButton::SetSelectedColor(COLORREF color)
{
	properties.selectedColor = color;
	if(IsWindow()) {
		Invalidate();
	}
}

COLORREF ColorButton::GetDefaultColor(void) const
{
	return properties.defaultColor;
}

void ColorButton::SetDefaultColor(COLORREF color)
{
	properties.defaultColor = color;
}

void ColorButton::SetMoreColorsText(LPCTSTR pText)
{
	if(pickerProperties.pMoreColorsText) {
		SECUREFREE(pickerProperties.pMoreColorsText);
	}
	if(pText) {
		int l = lstrlen(pText) + 1;
		pickerProperties.pMoreColorsText = reinterpret_cast<LPTSTR>(malloc(l * sizeof(TCHAR)));
		if(pickerProperties.pMoreColorsText) {
			lstrcpyn(pickerProperties.pMoreColorsText, pText, l);
		}
	}
}

void ColorButton::SetMoreColorsText(UINT resourceID)
{
	if(resourceID == 0) {
		SetMoreColorsText(reinterpret_cast<LPCTSTR>(NULL));
	} else {
		TCHAR buffer[256];
		if(LoadString(ModuleHelper::GetResourceInstance(), resourceID, buffer, sizeof(buffer) / sizeof(buffer[0]))) {
			SetMoreColorsText(buffer);
		}
	}
}

void ColorButton::SetDefaultColorText(LPCTSTR pText)
{
	if(pickerProperties.pDefaultColorText) {
		SECUREFREE(pickerProperties.pDefaultColorText);
	}
	if(pText) {
		int l = lstrlen(pText) + 1;
		pickerProperties.pDefaultColorText = reinterpret_cast<LPTSTR>(malloc(l * sizeof(TCHAR)));
		if(pickerProperties.pDefaultColorText) {
			lstrcpyn(pickerProperties.pDefaultColorText, pText, l);
		}
	}
}

void ColorButton::SetDefaultColorText(UINT resourceID)
{
	if(resourceID == 0) {
		SetDefaultColorText(reinterpret_cast<LPCTSTR>(NULL));
	} else {
		TCHAR buffer[256];
		if(LoadString(ModuleHelper::GetResourceInstance(), resourceID, buffer, sizeof(buffer) / sizeof(buffer[0]))) {
			SetDefaultColorText(buffer);
		}
	}
}

void ColorButton::SetTexts(UINT resourceIDDefaultColorText, UINT resourceIDMoreColorsText)
{
	SetDefaultColorText(resourceIDDefaultColorText);
	SetMoreColorsText(resourceIDMoreColorsText);
}

BOOL ColorButton::HasMoreColorsText(void) const
{
	return pickerProperties.pMoreColorsText && (lstrlen(pickerProperties.pMoreColorsText) > 0);
}

BOOL ColorButton::HasDefaultColorText(void) const
{
	return pickerProperties.pDefaultColorText && (lstrlen(pickerProperties.pDefaultColorText) > 0);
}

BOOL ColorButton::GetTrackSelection(void) const
{
	return properties.trackSelection;
}

void ColorButton::SetTrackSelection(BOOL trackSelect)
{
	properties.trackSelection = trackSelect;
}


LRESULT ColorButton::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(mouseStatus.enteredControl) {
		mouseStatus.enteredControl = FALSE;
		Invalidate();
	}
	return 0;
}

LRESULT ColorButton::OnMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(!mouseStatus.enteredControl) {
		mouseStatus.enteredControl = TRUE;
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = *this;
		trackingOptions.dwFlags = TME_LEAVE;
		TrackMouseEvent(&trackingOptions);

		Invalidate();
	}
	return 0;
}

LRESULT ColorButton::OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPDRAWITEMSTRUCT pDrawItemData = reinterpret_cast<LPDRAWITEMSTRUCT>(lParam);

	CDCHandle targetDC = pDrawItemData->hDC;
	WTL::CRect drawingRectangle = pDrawItemData->rcItem;

	UINT frameState;
	if(m_hTheme) {
		// draw the outer edge
		frameState = 0;
		if(((pDrawItemData->itemState & ODS_SELECTED) == ODS_SELECTED) || flags.colorPickerIsActive) {
			frameState |= PBS_PRESSED;
		}
		if((pDrawItemData->itemState & ODS_DISABLED) == ODS_DISABLED) {
			frameState |= PBS_DISABLED;
		}
		if(((pDrawItemData->itemState & ODS_HOTLIGHT) == ODS_HOTLIGHT) || mouseStatus.enteredControl) {
			frameState |= PBS_HOT;
		} else if((pDrawItemData->itemState & ODS_DEFAULT) == ODS_DEFAULT) {
			frameState |= PBS_DEFAULTED;
		}

		DrawThemeBackground(targetDC, BP_PUSHBUTTON, frameState, &drawingRectangle, NULL);
		GetThemeBackgroundContentRect(targetDC, BP_PUSHBUTTON, frameState, &drawingRectangle, &drawingRectangle);
	} else {
		// draw the outer edge
		frameState = DFCS_BUTTONPUSH | DFCS_ADJUSTRECT;
		if(((pDrawItemData->itemState & ODS_SELECTED) == ODS_SELECTED) || flags.colorPickerIsActive) {
			frameState |= DFCS_PUSHED;
		}
		if((pDrawItemData->itemState & ODS_DISABLED) == ODS_DISABLED) {
			frameState |= DFCS_INACTIVE;
		}

		targetDC.DrawFrameControl(&drawingRectangle, DFC_BUTTON, frameState);

		// adjust the position if we are selected (gives a 3d look)
		if(((pDrawItemData->itemState & ODS_SELECTED) == ODS_SELECTED) || flags.colorPickerIsActive) {
			drawingRectangle.OffsetRect(1, 1);
		}
	}

	// draw the focus rectangle
	if(((pDrawItemData->itemState & ODS_FOCUS) == ODS_FOCUS) || flags.colorPickerIsActive) {
		WTL::CRect focusRectangle(drawingRectangle.left, drawingRectangle.top, drawingRectangle.right - 1, drawingRectangle.bottom);
		targetDC.DrawFocusRect(&focusRectangle);
	}
	drawingRectangle.InflateRect(-GetSystemMetrics(SM_CXEDGE), -GetSystemMetrics(SM_CYEDGE));

	// draw the arrow
	WTL::CRect arrowRectangle;
	arrowRectangle.left = drawingRectangle.right - ARROWWIDTH - GetSystemMetrics(SM_CXEDGE) / 2;
	arrowRectangle.top = (drawingRectangle.bottom + drawingRectangle.top) / 2 - ARROWHEIGHT / 2;
	arrowRectangle.right = arrowRectangle.left + ARROWWIDTH;
	arrowRectangle.bottom = (drawingRectangle.bottom + drawingRectangle.top) / 2 + ARROWHEIGHT / 2;

	DrawArrow(targetDC, arrowRectangle, 0, (pDrawItemData->itemState & ODS_DISABLED) ? GetSysColor(COLOR_GRAYTEXT) : GetSysColor(COLOR_BTNTEXT));
	drawingRectangle.right = arrowRectangle.left - GetSystemMetrics(SM_CXEDGE) / 2;

	// draw the separator
	targetDC.DrawEdge(&drawingRectangle, EDGE_ETCHED, BF_RIGHT);
	drawingRectangle.right -= (GetSystemMetrics(SM_CXEDGE) * 2) + 1;

	// draw the color rectangle
	if((pDrawItemData->itemState & ODS_DISABLED) == 0) {
		if(properties.selectedColor == CLR_DEFAULT && properties.defaultColor == CLR_DEFAULT) {
			LPCTSTR pCellText = pickerProperties.pDefaultColorText;
			if(pCellText) {
				if(m_hTheme) {
					DrawThemeText(targetDC, BP_PUSHBUTTON, frameState, CT2CW(pCellText), -1, DT_CENTER | DT_VCENTER | DT_SINGLELINE, 0, &drawingRectangle);
				} else {
					HFONT hPreviousFont = targetDC.SelectFont(pickerProperties.font);
					COLORREF previousTextColor = targetDC.SetTextColor(GetSysColor(COLOR_BTNTEXT));
					int previousBkMode = targetDC.SetBkMode(TRANSPARENT);

					size_t l = 0;
					ATLVERIFY(SUCCEEDED(StringCchLength(pCellText, STRSAFE_MAX_CCH, &l)));
					targetDC.DrawText(pCellText, l, &drawingRectangle, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

					targetDC.SetBkMode(previousBkMode);
					targetDC.SetTextColor(previousTextColor);
					targetDC.SelectFont(hPreviousFont);
				}
			}
		} else {
			targetDC.SetBkColor((properties.selectedColor == CLR_DEFAULT) ? properties.defaultColor : properties.selectedColor);
			targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &drawingRectangle, NULL, 0, NULL);
		}
		targetDC.FrameRect(&drawingRectangle, reinterpret_cast<HBRUSH>(GetStockObject(BLACK_BRUSH)));
	}

	return TRUE;
}

LRESULT ColorButton::OnPickerPaint(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	CPaintDC targetDC(dropDownWindow);

	/* Draw a raised window edge. The extended window style WS_EX_WINDOWEDGE is supposed to do this, but
	   for some reason isn't. */

	WTL::CRect clientRectangle;
	dropDownWindow.GetClientRect(&clientRectangle);
	if(pickerProperties.has3DBorder) {
		CPen pen;
		pen.CreatePen(PS_SOLID, 0, GetSysColor(COLOR_GRAYTEXT));
		HPEN hPreviousPen = targetDC.SelectPen(pen);
		targetDC.Rectangle(&clientRectangle);
		targetDC.SelectPen(hPreviousPen);
	} else {
		targetDC.DrawEdge(&clientRectangle, EDGE_RAISED, BF_RECT);
	}

	// draw the "Automatic" text
	if(HasDefaultColorText()) {
		DrawPickerCell(targetDC, DEFAULTCOLORCELLINDEX);
	}

	// draw the color cells
	for(int i = 0; i < pickerProperties.numberOfPredefinedColors; ++i) {
		DrawPickerCell(targetDC, i);
	}

	// draw the "More Colors" text
	if(HasMoreColorsText()) {
		DrawPickerCell(targetDC, MORECOLORSCELLINDEX);
	}

	return 0;
}

LRESULT ColorButton::OnPickerPaletteChanged(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(reinterpret_cast<HWND>(wParam) != *this) {
		Invalidate();
	}
	return lr;
}

LRESULT ColorButton::OnPickerQueryNewPalette(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	Invalidate();
	return TRUE;
}

LRESULT ColorButton::OnPickerKeyDown(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	// get the offset for movement
	int offset = 0;
	switch(wParam) {
		case VK_DOWN:
			offset = pickerProperties.numberOfColumns;
			break;
		case VK_UP:
			offset = -pickerProperties.numberOfColumns;
			break;
		case VK_RIGHT:
			offset = 1;
			break;
		case VK_LEFT:
			offset = -1;
			break;
		case VK_ESCAPE:
			pickerProperties.hoveredColor = properties.selectedColor;
			EndPickerSelection(TRUE);
			break;
		case VK_RETURN:
		case VK_SPACE:
			if(pickerProperties.hoveredColorCell == INVALIDCOLORCELLINDEX) {
				pickerProperties.hoveredColor = properties.selectedColor;
			}
			EndPickerSelection(pickerProperties.hoveredColorCell == INVALIDCOLORCELLINDEX);
			break;
	}

	if(offset) {
		// compute a new position based on our current position
		int newSelection = 0;
		if(pickerProperties.hoveredColorCell == INVALIDCOLORCELLINDEX) {
			newSelection = offset > 0 ? DEFAULTCOLORCELLINDEX : MORECOLORSCELLINDEX;
		} else if(pickerProperties.hoveredColorCell == DEFAULTCOLORCELLINDEX) {
			newSelection = offset > 0 ? 0 : MORECOLORSCELLINDEX;
		} else if(pickerProperties.hoveredColorCell == MORECOLORSCELLINDEX) {
			newSelection = offset > 0 ? DEFAULTCOLORCELLINDEX : pickerProperties.numberOfPredefinedColors - 1;
		} else {
			newSelection = pickerProperties.hoveredColorCell + offset;
			if(newSelection < 0) {
				newSelection = DEFAULTCOLORCELLINDEX;
			} else if(newSelection >= pickerProperties.numberOfPredefinedColors) {
				newSelection = MORECOLORSCELLINDEX;
			}
		}

		// ensure we don't select cells that we don't have
		while(TRUE) {
			if((newSelection == DEFAULTCOLORCELLINDEX) && !HasDefaultColorText()) {
				newSelection = offset > 0 ? 0 : MORECOLORSCELLINDEX;
			} else if((newSelection == MORECOLORSCELLINDEX) && !HasMoreColorsText()) {
				newSelection = offset > 0 ? DEFAULTCOLORCELLINDEX : pickerProperties.numberOfPredefinedColors - 1;
			} else {
				break;
			}
		}

		// apply the new selection
		ChangePickerSelection(newSelection);
	}

	return 0;
}

LRESULT ColorButton::OnPickerLButtonUp(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// do a hit test
	WTL::CPoint mousePosition(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	int newSelection = PickerHitTest(mousePosition);

	if(newSelection != pickerProperties.hoveredColorCell) {
		ChangePickerSelection(newSelection);
	}
	EndPickerSelection(newSelection == INVALIDCOLORCELLINDEX);

	return 0;
}

LRESULT ColorButton::OnPickerMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// do a hit test
	WTL::CPoint mousePosition(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	int newSelection = PickerHitTest(mousePosition);

	if(newSelection != pickerProperties.hoveredColorCell) {
		ChangePickerSelection(newSelection);
	}
	return 0;
}

LRESULT ColorButton::OnButtonClicked(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/)
{
	// save the current color for future reference
	COLORREF previousColor = properties.selectedColor;

	// display the popup
	BOOL cancelled = !PopupPickerWindow();
	if(cancelled) {
		// if we are tracking, restore the old selection
		if(properties.trackSelection) {
			if(previousColor != properties.selectedColor) {
				properties.selectedColor = previousColor;
				SendNotification(CPN_SELCHANGE, properties.selectedColor, TRUE);
			}
		}
		SendNotification(CPN_CLOSEUP, properties.selectedColor, TRUE);
		SendNotification(CPN_SELENDCANCEL, properties.selectedColor, TRUE);
	} else {
		if(previousColor != properties.selectedColor) {
			SendNotification(CPN_SELCHANGE, properties.selectedColor, TRUE);
		}
		SendNotification(CPN_CLOSEUP, properties.selectedColor, TRUE);
		SendNotification(CPN_SELENDOK, properties.selectedColor, TRUE);
	}

	Invalidate();
	return 0;
}


BOOL ColorButton::PopupPickerWindow(void)
{
	BOOL cancelled = TRUE;

	// mark the button as active and invalidate button to force a redraw
	flags.colorPickerIsActive = TRUE;
	Invalidate();

	// send the drop down notification to the parent
	SendNotification(CPN_DROPDOWN, properties.selectedColor, TRUE);

	// initialize some sizing parameters
	ATLASSERT((sizeof(predefinedColors) / sizeof(ColorTableEntry)) <= MAXNUMBEROFCOLORCELLS);
	pickerProperties.SetNumberOfPredefinedColors(min(sizeof(predefinedColors) / sizeof(ColorTableEntry), MAXNUMBEROFCOLORCELLS));

	// initialize our state
	pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX;
	pickerProperties.selectedColorCell = INVALIDCOLORCELLINDEX;
	pickerProperties.hoveredColor = properties.selectedColor;

	/*OSVERSIONINFO windowsVersion = {0};
	windowsVersion.dwOSVersionInfoSize = sizeof(windowsVersion);
	GetVersionEx(&windowsVersion);
	BOOL usingWinXP = (windowsVersion.dwPlatformId == VER_PLATFORM_WIN32_NT) && ((windowsVersion.dwMajorVersion > 5) || ((windowsVersion.dwMajorVersion == 5) && (windowsVersion.dwMinorVersion >= 1)));*/
	BOOL usingWinXP = TRUE;

	// register the window class used for the picker
	WNDCLASSEX windowClass = {0};
	windowClass.cbSize = sizeof(WNDCLASSEX);
	windowClass.style = CS_CLASSDC | CS_SAVEBITS | CS_HREDRAW | CS_VREDRAW;
	windowClass.lpfnWndProc = CContainedWindow::StartWindowProc;
	windowClass.cbClsExtra = 0;
	windowClass.cbWndExtra = 0;
	windowClass.hInstance = ModuleHelper::GetModuleInstance();
	windowClass.hIcon = NULL;
	windowClass.hCursor = LoadCursor(NULL, IDC_ARROW);
	windowClass.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_MENU + 1);
	windowClass.lpszMenuName = NULL;
	windowClass.lpszClassName = TEXT("ColorPicker");
	windowClass.hIconSm = NULL;
	if(usingWinXP) {
		BOOL dropShadow = FALSE;
		SystemParametersInfo(SPI_GETDROPSHADOW, 0, &dropShadow, FALSE);
		if(dropShadow) {
			windowClass.style |= CS_DROPSHADOW;
		}
	}
	ATOM wndClassAtom = RegisterClassEx(&windowClass);

	// create the window
	WTL::CRect buttonRectangle;
	GetWindowRect(&buttonRectangle);
	ModuleHelper::AddCreateWndData(&dropDownWindow.m_thunk.cd, &dropDownWindow);
	dropDownWindow.m_hWnd = CreateWindowEx(0, MAKEINTATOM(wndClassAtom), TEXT(""), WS_POPUP, buttonRectangle.left, buttonRectangle.bottom, 100, 100, GetParent(), NULL, ModuleHelper::GetModuleInstance(), NULL);

	if(dropDownWindow) {
		// set the window size
		SetPickerWindowSize();

		// create the tooltips
		CToolTipCtrl toolTipControl;
		SetupPickerToolTips(toolTipControl);

		// find which cell (if any) corresponds to the initial color
		pickerProperties.selectedColorCell = FindPickerCellFromColor(properties.selectedColor);

		// make visible
		dropDownWindow.ShowWindow(SW_SHOWNA);

		// purge the message queue of paints
		MSG msg = {0};
		while(PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_NOREMOVE)) {
			if(!GetMessage(&msg, NULL, WM_PAINT, WM_PAINT)) {
				cancelled = TRUE;
				goto ret;
			}
			DispatchMessage(&msg);
		}

		// set capture to the window which received this message
		dropDownWindow.SetCapture();
		ATLASSERT(dropDownWindow == GetCapture());

		// get messages until capture lost or cancelled/accepted
		while(dropDownWindow == GetCapture()) {
			if(!GetMessage(&msg, NULL, 0, 0)) {
				PostQuitMessage(msg.wParam);
				break;
			}

			toolTipControl.RelayEvent(&msg);
			switch(msg.message) {
				case WM_LBUTTONUP: {
					BOOL wasHandled = TRUE;
					OnPickerLButtonUp(msg.message, msg.wParam, msg.lParam, wasHandled);
					break;
				}
				case WM_MOUSEMOVE: {
					BOOL wasHandled = TRUE;
					OnPickerMouseMove(msg.message, msg.wParam, msg.lParam, wasHandled);
					break;
				}
				case WM_KEYUP:
					break;
				case WM_KEYDOWN: {
					BOOL wasHandled = TRUE;
					OnPickerKeyDown(msg.message, msg.wParam, msg.lParam, wasHandled);
					break;
				}
				case WM_RBUTTONDOWN:
					ReleaseCapture();
					pickerProperties.cancelled = TRUE;
					break;
				default:
					DispatchMessage(&msg);
					break;
			}
		}
		ReleaseCapture();
		cancelled = pickerProperties.cancelled;

		// destroy the window
		toolTipControl.DestroyWindow();
		dropDownWindow.DestroyWindow();

		// if needed, show the color selection dialog
		if(!cancelled) {
			if(pickerProperties.hoveredColorCell == MORECOLORSCELLINDEX) {
				CColorDialog dlg(properties.selectedColor, CC_FULLOPEN | CC_ANYCOLOR, *this);
				if(dlg.DoModal() == IDOK) {
					properties.selectedColor = dlg.GetColor();
				} else {
					cancelled = TRUE;
				}
			} else {
				properties.selectedColor = pickerProperties.hoveredColor;
			}
		}
	}

ret:
	// unregister our class
	UnregisterClass(MAKEINTATOM(wndClassAtom), ModuleHelper::GetModuleInstance());

	flags.colorPickerIsActive = FALSE;
	return !cancelled;
}

void ColorButton::SetPickerWindowSize(void)
{
	SIZE textSize = {0, 0};

	// if we are showing a "More Colors" or "Automatic" text area, get the font and text size
	if(HasMoreColorsText() || HasDefaultColorText()) {
		CClientDC clientDC(dropDownWindow);
		HFONT hPreviousFont = clientDC.SelectFont(pickerProperties.font);

		// get the size of the "More Colors" text
		if(HasMoreColorsText()) {
			size_t l = 0;
			ATLVERIFY(SUCCEEDED(StringCchLength(pickerProperties.pMoreColorsText, STRSAFE_MAX_CCH, &l)));
			clientDC.GetTextExtent(pickerProperties.pMoreColorsText, l, &textSize);
		}

		// get the size of the "Automatic" text
		if(HasDefaultColorText()) {
			SIZE sz = {0};
			size_t l = 0;
			ATLVERIFY(SUCCEEDED(StringCchLength(pickerProperties.pDefaultColorText, STRSAFE_MAX_CCH, &l)));
			clientDC.GetTextExtent(pickerProperties.pDefaultColorText, l, &sz);
			textSize.cx = max(textSize.cx, sz.cx);
			textSize.cy = max(textSize.cy, sz.cy);
		}
		clientDC.SelectFont(hPreviousFont);

		// commpute the final size
		textSize.cx += 2 * (outerTextCellBorderSize.cx + innerTextCellBorderSize.cx);
		textSize.cy += 2 * (outerTextCellBorderSize.cy + innerTextCellBorderSize.cy);
	}

	// initiailize our box size
	ATLASSERT(innerColorCellBorderSize.cx == innerColorCellBorderSize.cy);
	ATLASSERT(outerColorCellBorderSize.cx == outerColorCellBorderSize.cy);
	pickerProperties.colorCellSize.cx = colorCellContentSize.cx + (innerColorCellBorderSize.cx + outerColorCellBorderSize.cx) * 2;
	pickerProperties.colorCellSize.cy = colorCellContentSize.cy + (innerColorCellBorderSize.cy + outerColorCellBorderSize.cy) * 2;

	// compute the minimum width
	int totalWidth = pickerProperties.numberOfColumns * pickerProperties.colorCellSize.cx;
	int minimumWidth = max(totalWidth, textSize.cx);

	// create the rectangle for the "Automatic" text
	pickerProperties.defaultColorTextBoundingRectangle = WTL::CRect(WTL::CPoint(0, 0), WTL::CSize(minimumWidth, HasDefaultColorText() ? textSize.cy : 0));

	// initialize the color box rectangle
	pickerProperties.colorGridBoundingRectangle = WTL::CRect(WTL::CPoint((minimumWidth - totalWidth) / 2, pickerProperties.defaultColorTextBoundingRectangle.bottom), WTL::CSize(totalWidth, pickerProperties.numberOfRows * pickerProperties.colorCellSize.cy));

	// create the rectangle for the "More Colors" text
	pickerProperties.moreColorsTextBoundingRectangle = WTL::CRect(WTL::CPoint(0, pickerProperties.colorGridBoundingRectangle.bottom), WTL::CSize(minimumWidth, HasMoreColorsText() ? textSize.cy : 0));

	// get the current window position, and set the new size
	WTL::CRect windowRectangle(pickerProperties.defaultColorTextBoundingRectangle.TopLeft(), pickerProperties.moreColorsTextBoundingRectangle.BottomRight());
	WTL::CRect rc;
	dropDownWindow.GetWindowRect(&rc);
	windowRectangle.OffsetRect(rc.TopLeft());

	// adjust the rects for the border
	windowRectangle.right += pickerProperties.margins.left + pickerProperties.margins.right;
	windowRectangle.bottom += pickerProperties.margins.top + pickerProperties.margins.bottom;
	pickerProperties.defaultColorTextBoundingRectangle.OffsetRect(pickerProperties.margins.left, pickerProperties.margins.top);
	pickerProperties.colorGridBoundingRectangle.OffsetRect(pickerProperties.margins.left, pickerProperties.margins.top);
	pickerProperties.moreColorsTextBoundingRectangle.OffsetRect(pickerProperties.margins.left, pickerProperties.margins.top);

	// get the screen rectangle
	WTL::CRect screenRectangle(WTL::CPoint(0, 0), WTL::CSize(GetSystemMetrics(SM_CXSCREEN), GetSystemMetrics(SM_CYSCREEN)));
	HMODULE hUser32 = GetModuleHandle(TEXT("user32.dll"));
	if(hUser32) {
		typedef HMONITOR WINAPI MonitorFromWindowFn(HWND, DWORD);
		typedef BOOL WINAPI GetMonitorInfoFn(HMONITOR, LPMONITORINFO);
		MonitorFromWindowFn* pMonitorFromWindow = reinterpret_cast<MonitorFromWindowFn*>(GetProcAddress(hUser32, "MonitorFromWindow"));
		#ifdef UNICODE
			GetMonitorInfoFn* pGetMonitorInfo = reinterpret_cast<GetMonitorInfoFn*>(GetProcAddress(hUser32, "GetMonitorInfoW"));
		#else
			GetMonitorInfoFn* pGetMonitorInfo = reinterpret_cast<GetMonitorInfoFn*>(GetProcAddress(hUser32, "GetMonitorInfoA"));
		#endif
		if(pMonitorFromWindow && pGetMonitorInfo) {
			MONITORINFO monitorInfo = {0};
			HMONITOR hMonitor = pMonitorFromWindow(*this, MONITOR_DEFAULTTONEAREST);
			monitorInfo.cbSize = sizeof(monitorInfo);
			pGetMonitorInfo(hMonitor, &monitorInfo);
			screenRectangle = monitorInfo.rcWork;
		}
	}

	// need to check it'll fit on screen: Too far right?
	if(windowRectangle.right > screenRectangle.right) {
		windowRectangle.OffsetRect(screenRectangle.right - windowRectangle.right, 0);
	}
	// Too far left?
	if(windowRectangle.left < screenRectangle.left) {
		windowRectangle.OffsetRect(screenRectangle.left - windowRectangle.left, 0);
	}

	// Bottom falling out of screen?  If so, move the whole popup above the parent window.
	if(windowRectangle.bottom > screenRectangle.bottom) {
		WTL::CRect parentWindowRectangle;
		GetWindowRect(&parentWindowRectangle);
		windowRectangle.OffsetRect(0, -((parentWindowRectangle.bottom - parentWindowRectangle.top) + (windowRectangle.bottom - windowRectangle.top)));
	}

	// set the window size and position
	dropDownWindow.MoveWindow(&windowRectangle, TRUE);
}

void ColorButton::SetupPickerToolTips(CToolTipCtrl& toolTipControl)
{
	if(!toolTipControl.Create(dropDownWindow)) {
		return;
	}

	// add a tool for each cell
	for(int i = 0; i < pickerProperties.numberOfPredefinedColors; ++i) {
		WTL::CRect cellRectangle;
		if(!GetPickerCellRectangle(i, &cellRectangle)) {
			continue;
		}
		toolTipControl.AddTool(dropDownWindow, predefinedColors[i].pName, &cellRectangle, 1);
	}
}

BOOL ColorButton::GetPickerCellRectangle(int cellIndex, LPRECT pRectangle) const
{
	if(cellIndex == MORECOLORSCELLINDEX) {
		*pRectangle = pickerProperties.moreColorsTextBoundingRectangle;
		return TRUE;
	} else if(cellIndex == DEFAULTCOLORCELLINDEX) {
		*pRectangle = pickerProperties.defaultColorTextBoundingRectangle;
		return TRUE;
	} else if((cellIndex < 0) || (cellIndex >= pickerProperties.numberOfPredefinedColors)) {
		return FALSE;
	} else {
		pRectangle->left = (cellIndex % pickerProperties.numberOfColumns) * pickerProperties.colorCellSize.cx + pickerProperties.colorGridBoundingRectangle.left;
		pRectangle->top  = (cellIndex / pickerProperties.numberOfColumns) * pickerProperties.colorCellSize.cy + pickerProperties.colorGridBoundingRectangle.top;
		pRectangle->right = pRectangle->left + pickerProperties.colorCellSize.cx;
		pRectangle->bottom = pRectangle->top + pickerProperties.colorCellSize.cy;
		return TRUE;
	}
}

int ColorButton::FindPickerCellFromColor(COLORREF color)
{
	if((color == CLR_DEFAULT) && HasDefaultColorText()) {
		return DEFAULTCOLORCELLINDEX;
	}

	for(int i = 0; i < pickerProperties.numberOfPredefinedColors; ++i) {
		if(predefinedColors[i].color == color) {
			return i;
		}
	}

	if(HasMoreColorsText()) {
		return MORECOLORSCELLINDEX;
	} else {
		return INVALIDCOLORCELLINDEX;
	}
}

int ColorButton::PickerHitTest(const POINT& point)
{
	if(pickerProperties.moreColorsTextBoundingRectangle.PtInRect(point)) {
		return MORECOLORSCELLINDEX;
	} else if(pickerProperties.defaultColorTextBoundingRectangle.PtInRect(point)) {
		return DEFAULTCOLORCELLINDEX;
	} else if(!pickerProperties.colorGridBoundingRectangle.PtInRect(point)) {
		return INVALIDCOLORCELLINDEX;
	} else {
		int row = (point.y - pickerProperties.colorGridBoundingRectangle.top) / pickerProperties.colorCellSize.cy;
		int col = (point.x - pickerProperties.colorGridBoundingRectangle.left) / pickerProperties.colorCellSize.cx;
		if((row < 0) || (row >= pickerProperties.numberOfRows) || (col < 0) || (col >= pickerProperties.numberOfColumns)) {
			return INVALIDCOLORCELLINDEX;
		}
		int cellIndex = row * pickerProperties.numberOfColumns + col;
		if(cellIndex >= pickerProperties.numberOfPredefinedColors) {
			return INVALIDCOLORCELLINDEX;
		}
		return cellIndex;
	}
}

void ColorButton::ChangePickerSelection(int cellIndex)
{
	CClientDC clientDC(dropDownWindow);

	if(cellIndex > pickerProperties.numberOfPredefinedColors) {
		cellIndex = MORECOLORSCELLINDEX;
	}

	// if the current selection is valid, redraw the old selection without it being selected
	if(((pickerProperties.hoveredColorCell >= 0) && (pickerProperties.hoveredColorCell < pickerProperties.numberOfPredefinedColors)) || (pickerProperties.hoveredColorCell == MORECOLORSCELLINDEX) || (pickerProperties.hoveredColorCell == DEFAULTCOLORCELLINDEX)) {
		int previousSelection = pickerProperties.hoveredColorCell;
		pickerProperties.hoveredColorCell = INVALIDCOLORCELLINDEX;
		DrawPickerCell(clientDC, previousSelection);
	}

	// set the current selection as row/col and draw (it will be drawn selected)
	pickerProperties.hoveredColorCell = cellIndex;
	DrawPickerCell(clientDC, pickerProperties.hoveredColorCell);

	// store the current color
	BOOL colorIsValid = TRUE;
	COLORREF color = RGB(0, 0, 0);
	if(pickerProperties.hoveredColorCell == MORECOLORSCELLINDEX) {
		color = properties.defaultColor;
	} else if(pickerProperties.hoveredColorCell == DEFAULTCOLORCELLINDEX) {
		color = pickerProperties.hoveredColor = CLR_DEFAULT;
	} else if(pickerProperties.hoveredColorCell == INVALIDCOLORCELLINDEX) {
		color = RGB(0, 0, 0);
		colorIsValid = FALSE;
	} else if(pickerProperties.hoveredColorCell >= 0 && pickerProperties.hoveredColorCell < sizeof(predefinedColors)) {
		color = pickerProperties.hoveredColor = predefinedColors[pickerProperties.hoveredColorCell].color;
	} else {
		ATLASSERT(FALSE);
	}

	// send notification
	if(properties.trackSelection) {
		if(colorIsValid) {
			properties.selectedColor = color;
		}
		Invalidate();
		SendNotification(CPN_SELCHANGE, properties.selectedColor, colorIsValid);
	}
}

void ColorButton::EndPickerSelection(BOOL cancelled)
{
	ReleaseCapture();
	pickerProperties.cancelled = cancelled;
}

void ColorButton::DrawPickerCell(CDC& targetDC, int cellIndex)
{
	// get the drawing rect
	WTL::CRect cellRectangle;
	if(!GetPickerCellRectangle(cellIndex, &cellRectangle)) {
		return;
	}

	// get the text and colors
	LPCTSTR pCellText = NULL;
	COLORREF color = NULL;
	SIZE outerBorderSize = {0};
	SIZE innerBorderSize = {0};
	if(cellIndex == MORECOLORSCELLINDEX) {
		pCellText = pickerProperties.pMoreColorsText;
		outerBorderSize = outerTextCellBorderSize;
		innerBorderSize = innerTextCellBorderSize;
	} else if(cellIndex == DEFAULTCOLORCELLINDEX) {
		pCellText = pickerProperties.pDefaultColorText;
		outerBorderSize = outerTextCellBorderSize;
		innerBorderSize = innerTextCellBorderSize;
	} else {
		color = predefinedColors[cellIndex].color;
		outerBorderSize = outerColorCellBorderSize;
		innerBorderSize = innerColorCellBorderSize;
	}

	// based on the selectons, get our colors
	COLORREF backgroundColor = NULL;
	COLORREF textColor = NULL;
	BOOL isSelected = FALSE;
	if(pickerProperties.hoveredColorCell == cellIndex) {
		isSelected = TRUE;
		backgroundColor = pickerProperties.hoverBackColor;
		textColor = pickerProperties.hoverTextColor;
	} else if(pickerProperties.selectedColorCell == cellIndex) {
		isSelected = TRUE;
		backgroundColor = pickerProperties.selectionBackColor;
		textColor = pickerProperties.selectionTextColor;
	} else {
		isSelected = FALSE;
		backgroundColor = pickerProperties.backColor;
		textColor = pickerProperties.selectionTextColor;
	}

	// select and realize the palette
	HPALETTE hPreviousPalette = NULL;
	if(!pCellText) {
		if(!pickerProperties.palette.IsNull() && ((targetDC.GetDeviceCaps(RASTERCAPS) & RC_PALETTE) != 0)) {
			hPreviousPalette = targetDC.SelectPalette(pickerProperties.palette, FALSE);
			targetDC.RealizePalette();
		}
	}

	if(isSelected) {
		// draw the background
		if((outerBorderSize.cx > 0) || (outerBorderSize.cy > 0)) {
			targetDC.SetBkColor(pickerProperties.backColor);
			targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &cellRectangle, NULL, 0, NULL);
			cellRectangle.InflateRect(-outerBorderSize.cx, -outerBorderSize.cy);
		}

		// draw the selection rectangle
		targetDC.SetBkColor(pickerProperties.highlightBorderColor);
		targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &cellRectangle, NULL, 0, NULL);
		cellRectangle.InflateRect(-1, -1);

		// draw the inner coloring
		targetDC.SetBkColor(backgroundColor);
		targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &cellRectangle, NULL, 0, NULL);
		cellRectangle.InflateRect(-(innerBorderSize.cx - 1), -(innerBorderSize.cy - 1));
	} else {
		// draw the background
		targetDC.SetBkColor(pickerProperties.backColor);
		targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &cellRectangle, NULL, 0, NULL);
		cellRectangle.InflateRect(-(outerBorderSize.cx + innerBorderSize.cx), -(outerBorderSize.cy + innerBorderSize.cy));
	}

	// draw the text or color
	if(pCellText) {
		HFONT hPreviousFont = targetDC.SelectFont(pickerProperties.font);
		targetDC.SetTextColor(textColor);
		targetDC.SetBkMode(TRANSPARENT);
		size_t l = 0;
		ATLVERIFY(SUCCEEDED(StringCchLength(pCellText, STRSAFE_MAX_CCH, &l)));
		targetDC.DrawText(pCellText, l, &cellRectangle, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
		targetDC.SelectFont(hPreviousFont);
	} else {
		targetDC.SetBkColor(GetSysColor(COLOR_3DSHADOW));
		targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &cellRectangle, NULL, 0, NULL);
		cellRectangle.InflateRect(-1, -1);
		ATLASSERT(cellIndex >= 0 && cellIndex < sizeof(predefinedColors));
		if(cellIndex >= 0 && cellIndex < sizeof(predefinedColors)) {
			targetDC.SetBkColor(predefinedColors[cellIndex].color);
		}
		targetDC.ExtTextOut(0, 0, ETO_OPAQUE, &cellRectangle, NULL, 0, NULL);
	}

	// restore the palette
	if((hPreviousPalette != NULL) && ((targetDC.GetDeviceCaps(RASTERCAPS) & RC_PALETTE) != 0)) {
		targetDC.SelectPalette(hPreviousPalette, FALSE);
	}
}

void ColorButton::DrawArrow(CDCHandle& targetDC, const RECT& boundingRectangle, int direction/* = 0*/, COLORREF color/* = RGB(0, 0, 0)*/)
{
	POINT corners[3];

	switch(direction) {
		case 0:
			// pointing downward
			corners[0].x = boundingRectangle.left;
			corners[0].y = boundingRectangle.top;
			corners[1].x = boundingRectangle.right;
			corners[1].y = boundingRectangle.top;
			corners[2].x = (boundingRectangle.left + boundingRectangle.right) / 2;
			corners[2].y = boundingRectangle.bottom;
			break;
		case 1:
			// pointing upward
			corners[0].x = boundingRectangle.left;
			corners[0].y = boundingRectangle.bottom;
			corners[1].x = boundingRectangle.right;
			corners[1].y = boundingRectangle.bottom;
			corners[2].x = (boundingRectangle.left + boundingRectangle.right) / 2;
			corners[2].y = boundingRectangle.top;
			break;
		case 2:
			// pointing to the left
			corners[0].x = boundingRectangle.right;
			corners[0].y = boundingRectangle.top;
			corners[1].x = boundingRectangle.right;
			corners[1].y = boundingRectangle.bottom;
			corners[2].x = boundingRectangle.left;
			corners[2].y = (boundingRectangle.top + boundingRectangle.bottom) / 2;
			break;
		case 3:
			// pointing to the right
			corners[0].x = boundingRectangle.left;
			corners[0].y = boundingRectangle.top;
			corners[1].x = boundingRectangle.left;
			corners[1].y = boundingRectangle.bottom;
			corners[2].x = boundingRectangle.right;
			corners[2].y = (boundingRectangle.top + boundingRectangle.bottom) / 2;
			break;
	}

	CBrush arrowBrush;
	arrowBrush.CreateSolidBrush(color);
	CPen arrowPen;
	arrowPen.CreatePen(PS_SOLID, 0, color);

	HBRUSH hPreviousBrush = targetDC.SelectBrush(arrowBrush);
	HPEN hPreviousPen = targetDC.SelectPen(arrowPen);

	targetDC.SetPolyFillMode(WINDING);
	targetDC.Polygon(corners, 3);

	targetDC.SelectBrush(hPreviousBrush);
	targetDC.SelectPen(hPreviousPen);
	return;
}

void ColorButton::SendNotification(UINT notifyCode, COLORREF color, BOOL colorIsValid)
{
	NMCOLORBUTTON notificationDetails = {0};

	notificationDetails.hdr.code = notifyCode;
	notificationDetails.hdr.hwndFrom = *this;
	notificationDetails.hdr.idFrom = GetDlgCtrlID();
	notificationDetails.colorIsValid = colorIsValid;
	notificationDetails.color = color;

	GetParent().SendMessage(WM_NOTIFY, GetDlgCtrlID(), reinterpret_cast<LPARAM>(&notificationDetails));
}