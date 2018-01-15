//////////////////////////////////////////////////////////////////////
/// \class ColorButton
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A button control that can be used as color picker</em>
///
/// This class provides an extended button control that has a dropdown window for picking a color.
///
/// \todo Auto-update PickerProperties::font, PickerProperties::margins, PickerProperties::has3DBorder,
///       PickerProperties::backColor, PickerProperties::highlightBorderColor,
///       PickerProperties::hoverBackColor, PickerProperties::hoverTextColor,
///       PickerProperties::selectionBackColor and PickerProperties::selectionTextColor if the system
///       settings are canged.
/// \todo Improve documentation.
///
/// \sa PanelProperties
//////////////////////////////////////////////////////////////////////


#pragma once

#include "stdafx.h"
#include "res/resource.h"
#include "helpers.h"


//////////////////////////////////////////////////////////////////////
/// \name Notification messages
///
//@{
/// \brief <em>The first notification code used by the control</em>
#define CPN_FIRST (0U - 1601U)
/// \brief <em>The selected color changed</em>
///
/// \sa SendNotification
#define CPN_SELCHANGE CPN_FIRST
/// \brief <em>The drop down window popped up</em>
///
/// \sa CPN_CLOSEUP, SendNotification
#define CPN_DROPDOWN (CPN_FIRST + 1)
/// \brief <em>The drop down window closed up</em>
///
/// \sa CPN_DROPDOWN, CPN_SELENDOK, CPN_SELENDCANCEL, SendNotification
#define CPN_CLOSEUP (CPN_FIRST + 2)
/// \brief <em>Color picking ended with success</em>
///
/// \sa CPN_CLOSEUP, CPN_SELENDCANCEL, SendNotification
#define CPN_SELENDOK (CPN_FIRST + 3)
/// \brief <em>Color picking was cancelled</em>
///
/// \sa CPN_CLOSEUP, CPN_SELENDOK, SendNotification
#define CPN_SELENDCANCEL (CPN_FIRST + 4)

/// \brief <em>Holds notification details</em>
typedef struct NMCOLORBUTTON
{
	/// \brief <em>Holds the notification detail's standard header</em>
	NMHDR hdr;
	/// \brief <em>If \c TRUE the color specified by \c color is valid; otherwise not</em>
	///
	/// \sa color
	BOOL colorIsValid;
	/// \brief <em>The color involved in this notification</em>
	///
	/// \sa colorIsValid
	COLORREF color;
} *LPNMCOLORBUTTON;
//@}
//////////////////////////////////////////////////////////////////////


class ColorButton :
    public CWindowImpl<ColorButton>,
    public CThemeImpl<ColorButton>
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ColorButton();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		BEGIN_MSG_MAP(ColorButton)
			CHAIN_MSG_MAP(CThemeImpl<ColorButton>)

			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)

			MESSAGE_HANDLER(OCM_DRAWITEM, OnReflectedDrawItem)

			REFLECTED_COMMAND_CODE_HANDLER(BN_CLICKED, OnButtonClicked)
			ALT_MSG_MAP(1)
			MESSAGE_HANDLER(WM_PAINT, OnPickerPaint)
			MESSAGE_HANDLER(WM_PALETTECHANGED, OnPickerPaletteChanged)
			MESSAGE_HANDLER(WM_QUERYNEWPALETTE, OnPickerQueryNewPalette)
		END_MSG_MAP()
	#endif

	/// \brief <em>Holds a predefined color's RGB value and name</em>
	typedef struct ColorTableEntry
	{
		/// \brief <em>The color's RGB value</em>
		COLORREF color;
		/// \brief <em>The color's name</em>
		LPCTSTR pName;
	} ColorTableEntry;

	//////////////////////////////////////////////////////////////////////
	/// \name Initialization
	///
	//@{
	/// \brief <em>Creates the button</em>
	///
	/// Call this to create the button.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] windowRect The button's size and position in relative coordinates.
	/// \param[in] lpWindowName The button's name.
	/// \param[in] windowStyle The button's window style bits.
	/// \param[in] extWindowStyle The button's extended window style bits.
	/// \param[in] controlID The button's control ID.
	/// \param[in] lpCreateParam Parameters that will be passed to the \c WM_CREATE handler.
	///
	/// \return The button's window handle if successful; otherwise \c NULL.
	HWND Create(HWND hWndParent, RECT& windowRect, LPCTSTR lpWindowName = NULL, DWORD windowStyle = 0, DWORD extWindowStyle = 0, UINT controlID = 0, LPVOID lpCreateParam = NULL);
	/// \brief <em>Creates the button</em>
	///
	/// \overload
	///
	/// Call this to create the button.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] lpWindowRect The button's window rectangle in relative coordinates.
	/// \param[in] lpWindowName The button's name.
	/// \param[in] windowStyle The button's window style bits.
	/// \param[in] extWindowStyle The button's extended window style bits.
	/// \param[in] hMenu Handle to The button's menu.
	/// \param[in] lpCreateParam Parameters that will be passed to the WM_CREATE handler.
	///
	/// \return The button's window handle if successful; otherwise \c NULL.
	HWND Create(HWND hWndParent, LPRECT lpWindowRect = NULL, LPCTSTR lpWindowName = NULL, DWORD windowStyle = 0, DWORD extWindowStyle = 0, HMENU hMenu = NULL, LPVOID lpCreateParam = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Subclasses an existing button</em>
	///
	/// Call this to subclass an existing button.
	///
	/// \param[in] hWnd The button's window handle.
	///
	/// \return \c TRUE if successful; otherwise \c FALSE.
	BOOL SubclassWindow(HWND hWnd);

	//////////////////////////////////////////////////////////////////////
	/// \name General operations
	///
	//@{
	/// \brief <em>Retrieves the currently selected color</em>
	///
	/// \return The currently selected color.
	///
	/// \sa SetSelectedColor, GetDefaultColor
	COLORREF GetSelectedColor(void) const;
	/// \brief <em>Sets the currently selected color</em>
	///
	/// \param[in] color The color to select.
	///
	/// \sa GetSelectedColor, SetDefaultColor
	void SetSelectedColor(COLORREF color);
	/// \brief <em>Retrieves the current default color</em>
	///
	/// \return The current default color.
	///
	/// \sa SetDefaultColor, GetSelectedColor
	COLORREF GetDefaultColor(void) const;
	/// \brief <em>Sets the current default color</em>
	///
	/// \param[in] color The color to make the default.
	///
	/// \sa GetDefaultColor, SetSelectedColor, HasMoreColorsText
	void SetDefaultColor(COLORREF color);
	/// \brief <em>Sets the text for the "More Colors" button</em>
	///
	/// \param[in] pText The text to set.
	///
	/// \sa SetDefaultColorText, SetTexts, HasMoreColorsText
	void SetMoreColorsText(LPCTSTR pText);
	/// \brief <em>Sets the text for the "More Colors" button</em>
	///
	/// \overload
	///
	/// \param[in] resourceID The resource ID of the text to set.
	///
	/// \sa SetDefaultColorText, SetTexts
	void SetMoreColorsText(UINT resourceID);
	/// \brief <em>Sets the text for the "Automatic" button</em>
	///
	/// \param[in] pText The text to set.
	///
	/// \sa SetMoreColorsText, SetTexts, HasDefaultColorText
	void SetDefaultColorText(LPCTSTR pText);
	/// \brief <em>Sets the text for the "Automatic" button</em>
	///
	/// \overload
	///
	/// \param[in] resourceID The resource ID of the text to set.
	///
	/// \sa SetMoreColorsText, SetTexts, HasDefaultColorText
	void SetDefaultColorText(UINT resourceID);
	/// \brief <em>Sets the texts for the "Automatic" and the "More Colors" buttons</em>
	///
	/// \param[in] resourceID The resource ID of the text to set.
	///
	/// \sa SetDefaultColorText, SetMoreColorsText
	void SetTexts(UINT resourceIDDefaultColorText, UINT resourceIDMoreColorsText);
	/// \brief <em>Retrieves whether we've a text for the "More Colors" button</em>
	///
	/// \return \c TRUE if we have a text for the "More Colors" button; otherwise \c FALSE.
	///
	/// \sa HasDefaultColorText
	BOOL HasMoreColorsText(void) const;
	/// \brief <em>Retrieves whether we've a text for the "Automatic" button</em>
	///
	/// \return \c TRUE if we have a text for the "Automatic" button; otherwise \c FALSE.
	///
	/// \sa HasMoreColorsText
	BOOL HasDefaultColorText(void) const;
	/// \brief <em>Retrieves the current setting of the track selection feature</em>
	///
	/// \return \c TRUE if track selection is activated; otherwise \c FALSE.
	///
	/// \sa SetTrackSelection
	BOOL GetTrackSelection(void) const;
	/// \brief <em>Activates or deactivates the track selection feature</em>
	///
	/// \param[in] trackSelect If \c TRUE, track selection is activated; otherwise not.
	///
	/// \sa GetTrackSelection
	void SetTrackSelection(BOOL trackSelect);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	/// \brief <em>Specifies the drop down arrow's width in pixles</em>
	///
	/// \sa ARROWHEIGHT
	static const int ARROWWIDTH = 4;
	/// \brief <em>Specifies the drop down arrow's height in pixles</em>
	///
	/// \sa ARROWWIDTH
	static const int ARROWHEIGHT = 2;

	/// \brief <em>Identifies the "Automatic" color cell</em>
	///
	/// \sa MORECOLORSCELLINDEX, INVALIDCOLORCELLINDEX
	static const int DEFAULTCOLORCELLINDEX = -3;
	/// \brief <em>Identifies the "More Colors" color cell</em>
	///
	/// \sa DEFAULTCOLORCELLINDEX, INVALIDCOLORCELLINDEX
	static const int MORECOLORSCELLINDEX = -2;
	/// \brief <em>Identifies an invalid color cell</em>
	///
	/// \sa DEFAULTCOLORCELLINDEX, MORECOLORSCELLINDEX
	static const int INVALIDCOLORCELLINDEX = -1;
	/// \brief <em>Specifies the maximum number of cells in the color grid</em>
	static const int MAXNUMBEROFCOLORCELLS = 100;

	//////////////////////////////////////////////////////////////////////
	/// \name Sizings
	///
	//@{
	/// \brief <em>Specifies the size of the inner border around a text cell</em>
	///
	/// \sa outerTextCellBorderSize, innerColorCellBorderSize
	static const WTL::CSize innerTextCellBorderSize;
	/// \brief <em>Specifies the size of the outer border around a text cell</em>
	///
	/// \sa innerTextCellBorderSize, outerColorCellBorderSize
	static const WTL::CSize outerTextCellBorderSize;
	/// \brief <em>Specifies the size of the inner border around a color cell</em>
	///
	/// \sa outerColorCellBorderSize, innerTextCellBorderSize, colorCellContentSize
	static const WTL::CSize innerColorCellBorderSize;
	/// \brief <em>Specifies the size of the outer border around a color cell</em>
	///
	/// \sa innerColorCellBorderSize, outerTextCellBorderSize, colorCellContentSize
	static const WTL::CSize outerColorCellBorderSize;
	/// \brief <em>Specifies the size of a color cell's content</em>
	///
	/// \sa innerColorCellBorderSize, outerColorCellBorderSize
	static const WTL::CSize colorCellContentSize;
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to redraw.
	///
	/// \sa OnMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to redraw.
	///
	/// \sa OnMouseLeave, OnPickerMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DRAWITEM handler</em>
	///
	/// Will be called if the button needs to be drawn.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms673316.aspx">WM_DRAWITEM</a>
	LRESULT OnReflectedDrawItem(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the popup window must be drawn.
	/// We use this handler to draw the color grid.
	///
	/// \sa OnPickerQueryNewPalette, OnPickerPaletteChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnPickerPaint(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_PALETTECHANGED handler</em>
	///
	/// Will be called after the focused window has realized its logical palette, thereby changing the
	/// system palette.
	/// We use this handler to redraw.
	///
	/// \sa OnPickerQueryNewPalette, OnPickerPaint,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms532653.aspx">WM_PALETTECHANGED</a>
	LRESULT OnPickerPaletteChanged(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_QUERYNEWPALETTE handler</em>
	///
	/// Will be called if we're about to receive the keyboard focus and should realize our logical palette.
	/// We use this handler to redraw.
	///
	/// \sa OnPickerPaletteChanged, OnPickerPaint,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms532654.aspx">WM_QUERYNEWPALETTE</a>
	LRESULT OnPickerQueryNewPalette(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the popup window has the focus.
	/// We use this handler to allow keyboard navigation within the color grid.
	///
	/// \sa OnMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnPickerKeyDown(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the popup window's client area.
	/// We use this handler to close the color picker.
	///
	/// \sa OnPickerMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnPickerLButtonUp(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the popup window's
	/// client area.
	/// We use this handler to redraw.
	///
	/// \sa OnMouseMove, OnPickerLButtonUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnPickerMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c BN_CLICKED handler</em>
	///
	/// Will be called if the button was clicked.
	/// We use this handler to popup the color picker.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms673572.aspx">BN_CLICKED</a>
	LRESULT OnButtonClicked(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Helper methods
	///
	//@{
	/// \brief <em>Displays the drop down color picker</em>
	///
	/// \return \c TRUE if the user selected a color; otherwise \c FALSE.
	///
	/// \sa SetPickerWindowSize
	BOOL PopupPickerWindow(void);
	/// \brief <em>Sizes the drop down color picker</em>
	///
	/// \sa PopupPickerWindow
	void SetPickerWindowSize(void);
	/// \brief <em>Creates and configures the color cell's tooltip control</em>
	///
	/// \param[in] toolTipControl The tooltip control to use.
	void SetupPickerToolTips(CToolTipCtrl& toolTipControl);
	/// \brief <em>Retrieves a drop down color picker cell's rectangle</em>
	///
	/// \param[in] cellIndex The cell for which to retrieve the rectangle.
	/// \param[out] pRectangle The cell's rectangle.
	///
	/// \return \c TRUE if the specified cell is valid; otherwise \c FALSE.
	///
	/// \sa PickerHitTest, DEFAULTCOLORCELLINDEX, MORECOLORSCELLINDEX
	BOOL GetPickerCellRectangle(int cellIndex, LPRECT pRectangle) const;
	/// \brief <em>Retrieves the specified color's corresponding cell index</em>
	///
	/// \param[in] color The color to search for.
	///
	/// \return The index of the color's corresponding cell.
	///
	/// \sa PickerHitTest
	int FindPickerCellFromColor(COLORREF color);
	/// \brief <em>Retrieves the index of the drop down color picker cell that the specified point lies in</em>
	///
	/// \param[in] point The point to retrieve the cell for.
	///
	/// \return The index of the drop down color picker cell that the specified point lies in.
	///
	/// \sa GetPickerCellRectangle
	int PickerHitTest(const POINT& point);
	/// \brief <em>Changes the drop down color picker's selected color</em>
	///
	/// \param[in] cellIndex The index of the new color to select.
	///
	/// \sa EndPickerSelection, DrawPickerCell
	void ChangePickerSelection(int cellIndex);
	/// \brief <em>Ends the color selection process</em>
	///
	/// \param[in] cancelled If \c TRUE, the user cancelled color selection; otherwise not.
	///
	/// \sa PopupPickerWindow
	void EndPickerSelection(BOOL cancelled);
	/// \brief <em>Draws the specified drop down color picker cell</em>
	///
	/// \param[in] targetDC The device context in which all drawing will take place.
	/// \param[in] cellIndex The index of the color cell to draw.
	void DrawPickerCell(CDC& targetDC, int cellIndex);
	/// \brief <em>Draws an arrow</em>
	///
	/// \param[in] targetDC The device context in which all drawing will take place.
	/// \param[in] boundingRectangle The arrow's bounding rectangle.
	/// \param[in] direction The arrow's direction. If 0, the drawn arrow points downward; if 1, it points
	///            upward; if 2, it points to the left; if 3, it points to the right.
	/// \param[in] color The arrow's color.
	static void DrawArrow(CDCHandle& targetDC, const RECT& boundingRectangle, int direction = 0, COLORREF color = RGB(0, 0, 0));
	/// \brief <em>Sends the parent window a notification</em>
	///
	/// \param[in] notifyCode The notification to send.
	/// \param[in] color The notification's involved color.
	/// \param[in] colorIsValid If \c TRUE, \c color specifies a valid color; otherwise not.
	///
	/// \sa CPN_SELCHANGE, CPN_DROPDOWN, CPN_CLOSEUP, CPN_SELENDOK, CPN_SELENDCANCEL, NMCOLORBUTTON
	void SendNotification(UINT notifyCode, COLORREF color, BOOL colorIsValid);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the button control's settings</em>
	///
	/// \sa pickerProperties
	struct Properties
	{
		/// \brief <em>Holds the default color's value</em>
		///
		/// \sa GetDefaultColor, SetDefaultColor, selectedColor
		COLORREF defaultColor;
		/// \brief <em>Holds the currently selected color's value</em>
		///
		/// \sa GetSelectedColor, SetSelectedColor, defaultColor
		COLORREF selectedColor;
		/// \brief <em>Holds whether the track selection feature is active or not</em>
		///
		/// \sa GetTrackSelection, SetTrackSelection
		UINT trackSelection : 1;

		Properties()
		{
			defaultColor = CLR_DEFAULT;
			selectedColor = CLR_DEFAULT;
			trackSelection = FALSE;
		}
	} properties;
	/// \brief <em>Wraps the drop down window</em>
	///
	/// \sa PopupPickerWindow
	CContainedWindow dropDownWindow;

	/// \brief <em>Holds the drop down window's settings</em>
	///
	/// \sa properties
	struct PickerProperties
	{
		/// \brief <em>Holds the "Automatic" cell's text</em>
		///
		/// \sa SetDefaultColorText, pMoreColorsText
		LPTSTR pDefaultColorText;
		/// \brief <em>Holds the "More Colors" cell's text</em>
		///
		/// \sa SetMoreColorsText, pDefaultColorText, moreColorsTextBoundingRectangle
		LPTSTR pMoreColorsText;
		/// \brief <em>Wraps the font used to draw text cells</em>
		///
		/// \sa pDefaultColorText, pMoreColorsText
		CFont font;
		/// \brief <em>Holds the number of the color grid's columns</em>
		///
		/// \sa numberOfRows, numberOfPredefinedColors, SetPickerWindowSize
		int numberOfColumns;
		/// \brief <em>Holds the number of the color grid's rows</em>
		///
		/// \sa numberOfColumns, numberOfPredefinedColors, SetPickerWindowSize
		int numberOfRows;
		/// \brief <em>Holds the number of the predefined colors</em>
		///
		/// \sa numberOfColumns, SetNumberOfPredefinedColors
		int numberOfPredefinedColors;
		/// \brief <em>Wraps the palette used to draw the drop down color picker</em>
		///
		/// \sa DrawPickerCell
		CPalette palette;
		/// \brief <em>Holds the value of the color that is currently marked for selection</em>
		///
		/// \sa ChangePickerSelection, hoveredColorCell
		COLORREF hoveredColor;
		/// \brief <em>Holds the cell index of the color that is currently marked for selection</em>
		///
		/// \sa ChangePickerSelection, hoveredColor
		int hoveredColorCell;
		/// \brief <em>Holds the currently selected color cell's index</em>
		///
		/// \sa ChangePickerSelection, hoveredColorCell
		int selectedColorCell;
		/// \brief <em>Holds the drop down color picker's margins</em>
		///
		/// \sa SetPickerWindowSize
		WTL::CRect margins;
		/// \brief <em>Holds the "More Colors" text's bounding rectangle</em>
		///
		/// \sa pMoreColorsText, SetPickerWindowSize
		WTL::CRect moreColorsTextBoundingRectangle;
		/// \brief <em>Holds the "Automatic" text's bounding rectangle</em>
		///
		/// \sa pDefaultColorText, SetPickerWindowSize
		WTL::CRect defaultColorTextBoundingRectangle;
		/// \brief <em>Holds the bounding rectangle around the predefined color cells</em>
		///
		/// \sa predefinedColors, SetPickerWindowSize
		WTL::CRect colorGridBoundingRectangle;
		/// \brief <em>Holds whether the drop down color picker has a 3D border or not</em>
		///
		/// \sa OnPickerPaint
		BOOL has3DBorder;
		/// \brief <em>Holds a color cell's size including inner and outer borders</em>
		///
		/// \sa SetPickerWindowSize
		WTL::CSize colorCellSize;
		/// \brief <em>Holds whether the drop down color picker was cancelled or not</em>
		///
		/// \sa PopupPickerWindow
		UINT cancelled : 1;
		/// \brief <em>Holds the drop down color picker's background color</em>
		///
		/// \sa DrawPickerCell
		COLORREF backColor;
		/// \brief <em>Holds the color of the border of a selected or hovered cell</em>
		///
		/// \sa DrawPickerCell, hoverBackColor, hoverTextColor
		COLORREF highlightBorderColor;
		/// \brief <em>Holds the background color of the hovered cell</em>
		///
		/// \sa DrawPickerCell, highlightBorderColor, hoverTextColor, selectionBackColor
		COLORREF hoverBackColor;
		/// \brief <em>Holds the text color of the hovered cell</em>
		///
		/// \sa DrawPickerCell, highlightBorderColor, hoverBackColor, selectionTextColor
		COLORREF hoverTextColor;
		/// \brief <em>Holds the background color of the hovered cell</em>
		///
		/// \sa DrawPickerCell, highlightBorderColor, selectionTextColor, hoverBackColor
		COLORREF selectionBackColor;
		/// \brief <em>Holds the text color of the hovered cell</em>
		///
		/// \sa DrawPickerCell, highlightBorderColor, selectionBackColor, hoverTextColor
		COLORREF selectionTextColor;

		PickerProperties()
		{
			pDefaultColorText = reinterpret_cast<LPTSTR>(malloc(10 * sizeof(TCHAR)));
			if(pDefaultColorText) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pDefaultColorText, 10, TEXT("Automatic"))));
			}
			pMoreColorsText = reinterpret_cast<LPTSTR>(malloc(15 * sizeof(TCHAR)));
			if(pMoreColorsText) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pMoreColorsText, 15, TEXT("More Colors..."))));
			}

			// create the font
			NONCLIENTMETRICS nonClientMetrics = {0};
			nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
			font.CreateFontIndirect(&nonClientMetrics.lfMessageFont);

			margins.left = GetSystemMetrics(SM_CXEDGE);
			margins.top = GetSystemMetrics(SM_CYEDGE);
			margins.right = GetSystemMetrics(SM_CXEDGE);
			margins.bottom = GetSystemMetrics(SM_CYEDGE);

			/*OSVERSIONINFO windowsVersion = {0};
			windowsVersion.dwOSVersionInfoSize = sizeof(windowsVersion);
			GetVersionEx(&windowsVersion);
			BOOL usingWinXP = (windowsVersion.dwPlatformId == VER_PLATFORM_WIN32_NT) && ((windowsVersion.dwMajorVersion > 5) || ((windowsVersion.dwMajorVersion == 5) && (windowsVersion.dwMinorVersion >= 1)));*/
			BOOL usingWinXP = TRUE;
			has3DBorder = FALSE;
			if(usingWinXP) {
				SystemParametersInfo(SPI_GETFLATMENU, 0, &has3DBorder, FALSE);
			}

			// get all the colors we need
			backColor = GetSysColor(COLOR_MENU);

			highlightBorderColor = GetSysColor(COLOR_HIGHLIGHT);
			if(usingWinXP) {
				hoverBackColor = GetSysColor(COLOR_MENUHILIGHT);
			} else {
				hoverBackColor = highlightBorderColor;
			}
			hoverTextColor = GetSysColor(COLOR_HIGHLIGHTTEXT);

			selectionTextColor = GetSysColor(COLOR_MENUTEXT);
			int alphaChannel = 48;
			selectionBackColor = RGB((LOBYTE(backColor) * (255 - alphaChannel) + LOBYTE(highlightBorderColor) * alphaChannel) >> 8, (LOBYTE(backColor >> 8) * (255 - alphaChannel) + LOBYTE(highlightBorderColor >> 8) * alphaChannel) >> 8, (LOBYTE(backColor >> 16) * (255 - alphaChannel) + LOBYTE(highlightBorderColor >> 16) * alphaChannel) >> 8);
		}

		~PickerProperties()
		{
			if(pDefaultColorText) {
				SECUREFREE(pDefaultColorText);
			}
			if(pMoreColorsText) {
				SECUREFREE(pMoreColorsText);
			}

			if(!font.IsNull()) {
				font.DeleteObject();
			}
			if(!palette.IsNull()) {
				palette.DeleteObject();
			}
		}

		/// \brief <em>Sets the number of predefined colors</em>
		///
		/// \param[in] number The new number of predefined colors.
		///
		/// \sa numberOfPredefinedColors, numberOfColumns, numberOfRows
		void SetNumberOfPredefinedColors(int number)
		{
			numberOfPredefinedColors = min(number, MAXNUMBEROFCOLORCELLS);
			// retrieve the number of columns and rows
			numberOfColumns = 8;
			numberOfRows = numberOfPredefinedColors / numberOfColumns;
			if((numberOfPredefinedColors % numberOfColumns) != 0) {
				++numberOfRows;
			}

			// create the palette
			if(!palette.IsNull()) {
				palette.DeleteObject();
			}

			PLOGPALETTE pLogPalette = static_cast<PLOGPALETTE>(HeapAlloc(GetProcessHeap(), 0, sizeof(LOGPALETTE) + MAXNUMBEROFCOLORCELLS * sizeof(PALETTEENTRY)));
			if(pLogPalette) {
				pLogPalette->palVersion = 0x300;
				pLogPalette->palNumEntries = static_cast<WORD>(numberOfPredefinedColors);

				for(int i = 0; i < numberOfPredefinedColors; ++i) {
					pLogPalette->palPalEntry[i].peRed = LOBYTE(predefinedColors[i].color);
					pLogPalette->palPalEntry[i].peGreen = LOBYTE(predefinedColors[i].color >> 8);
					pLogPalette->palPalEntry[i].peBlue = LOBYTE(predefinedColors[i].color >> 16);
					pLogPalette->palPalEntry[i].peFlags = 0;
				}
				palette.CreatePalette(pLogPalette);
				HeapFree(GetProcessHeap(), 0, pLogPalette);
			}
		}
	} pickerProperties;
	/// \brief <em>Holds the RGB values and names of the predefined colors</em>
	///
	/// \sa ColorTableEntry
	static ColorTableEntry predefinedColors[];

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the drop down color picker is visible</em>
		///
		/// \sa PopupPickerWindow
		UINT colorPickerIsActive : 1;

		Flags()
		{
			colorPickerIsActive = FALSE;
		}
	} flags;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
		/// \brief <em>If \c TRUE, we already called TrackMouseEvent to watch for the mouse leaving the client area</em>
		///
		/// \sa OnMouseMove, OnMouseLeave
		UINT enteredControl : 1;

		MouseStatus()
		{
			enteredControl = FALSE;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

private:
};