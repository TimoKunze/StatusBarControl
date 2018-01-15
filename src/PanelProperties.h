//////////////////////////////////////////////////////////////////////
/// \class PanelProperties
/// \author Timo "TimoSoft" Kunze
/// \brief <em>The control's "Panels" property page</em>
///
/// This class contains the control's "Panels" property page.
///
/// \sa StatusBar
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "StatBarU.h"
#else
	#include "StatBarA.h"
#endif
#include "StatusBar.h"
#include "ColorButton.h"


class ATL_NO_VTABLE PanelProperties :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<PanelProperties, &CLSID_PanelProperties>,
    public IPropertyPage2Impl<PanelProperties>,
    public CDialogImpl<PanelProperties>
{
public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	PanelProperties();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		enum { IDD = IDD_PANELPROPERTIES };

		DECLARE_REGISTRY_RESOURCEID(IDR_PANELPROPERTIES)

		BEGIN_COM_MAP(PanelProperties)
			COM_INTERFACE_ENTRY(IPropertyPage)
			//COM_INTERFACE_ENTRY2(IPropertyPage2, IPropertyPage)
		END_COM_MAP()

		BEGIN_MSG_MAP(PanelProperties)
			MESSAGE_HANDLER(WM_DRAWITEM, OnDrawItem)

			NOTIFY_CODE_HANDLER(CPN_SELCHANGE, OnColorButtonSelChangeNotification)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOA, OnToolTipGetDispInfoNotificationA)
			NOTIFY_CODE_HANDLER(TTN_GETDISPINFOW, OnToolTipGetDispInfoNotificationW)
			NOTIFY_CODE_HANDLER(UDN_DELTAPOS, OnUpDownDeltaPosNotification)

			COMMAND_ID_HANDLER(ID_LOADSETTINGS, OnLoadSettingsFromFile)
			COMMAND_ID_HANDLER(ID_SAVESETTINGS, OnSaveSettingsToFile)

			COMMAND_CODE_HANDLER(BN_CLICKED, OnButtonClicked)
			COMMAND_CODE_HANDLER(CBN_SELCHANGE, OnComboSelChange)
			COMMAND_CODE_HANDLER(EN_CHANGE, OnEditChange)
			COMMAND_CODE_HANDLER(EN_KILLFOCUS, OnEditKillFocus)

			REFLECT_NOTIFICATIONS()
		END_MSG_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPropertyPage
	///
	//@{
	/// \brief <em>Creates the property page's dialog box</em>
	///
	/// Creates the property page's dialog box window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetObjects,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682250.aspx">IPropertyPage::Activate</a>
	virtual HRESULT STDMETHODCALLTYPE Activate(HWND hWndParent, LPCRECT pRect, BOOL modal);
	/// \brief <em>Applies the current values to the controls</em>
	///
	/// Applies the current values to the underlying objects associated with the property page.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SetObjects,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms691284.aspx">IPropertyPage::Apply</a>
	virtual HRESULT STDMETHODCALLTYPE Apply(void);
	/// \brief <em>Provides the controls this property page is associated to</em>
	///
	/// \param[in] objects The number of objects defined by the \c ppControls array.
	/// \param[in] ppControls An array of \c IUnknown pointers each identifying an associated control.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Activate, Apply,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms678529.aspx">IPropertyPage::SetObjects</a>
	virtual HRESULT STDMETHODCALLTYPE SetObjects(ULONG objects, IUnknown** ppControls);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_DRAWITEM handler</em>
	///
	/// Will be called if a combobox item needs to be drawn in a ownerdrawn combobox. Used to draw the
	/// items in the \c panelForeColorCombo combobox.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms673316.aspx">WM_DRAWITEM</a>
	LRESULT OnDrawItem(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c CPN_SELCHANGE handler</em>
	///
	/// Will be called if the \c panelForeColorButton control's selection was changed.
	/// We use this handler to set the \c dirty flag.
	LRESULT OnColorButtonSelChangeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called if a tooltip control requests the text to display. We use this handler to display
	/// tooltips for the toolbar buttons.
	///
	/// \sa OnToolTipGetDispInfoNotificationW,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called if a tooltip control requests the text to display. We use this handler to display
	/// tooltips for the toolbar buttons.
	///
	/// \sa OnToolTipGetDispInfoNotificationA,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms650474.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c UDN_DELTAPOS handler</em>
	///
	/// Will be called if the \c panelIndexUpDown control's value is about to change.
	/// We use this handler to save the current selected panel's properties and to load those of the new
	/// selected panel.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms649979.aspx">UDN_DELTAPOS</a>
	LRESULT OnUpDownDeltaPosNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c ID_LOADSETTINGS handler</em>
	///
	/// Will be called if the \c ID_LOADSETTINGS command was initiated.
	/// We use this handler to execute the command.
	LRESULT OnLoadSettingsFromFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c ID_SAVESETTINGS handler</em>
	///
	/// Will be called if the \c ID_SAVESETTINGS command was initiated.
	/// We use this handler to execute the command.
	LRESULT OnSaveSettingsToFile(WORD /*notifyCode*/, WORD /*controlID*/, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c BN_CLICKED handler</em>
	///
	/// Will be called if a button was clicked.
	/// We use this handler to execute the corresponding action.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms673572.aspx">BN_CLICKED</a>
	LRESULT OnButtonClicked(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& wasHandled);
	/// \brief <em>\c CBN_SELCHANGE handler</em>
	///
	/// Will be called if a combo box's selection was changed.
	/// We use this handler to set the \c dirty flag.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms673136.aspx">CBN_SELCHANGE</a>
	LRESULT OnComboSelChange(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c EN_CHANGE handler</em>
	///
	/// Will be called if an edit control's text was changed.
	/// We use this handler to set the \c dirty flag.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672110.aspx">EN_CHANGE</a>
	LRESULT OnEditChange(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c EN_KILLFOCUS handler</em>
	///
	/// Will be called if an edit control lost the keyboard focus.
	/// We use this handler for the \c panelIndexEdit control to change the selected panel.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms672114.aspx">EN_KILLFOCUS</a>
	LRESULT OnEditKillFocus(WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates an \c OLE_COLOR type color into a \c COLORREF type color</em>
	///
	/// \param[in] oleColor The \c OLE_COLOR type color to translate.
	/// \param[in] hPalette The color palette to use.
	///
	/// \return The translated \c COLORREF type color.
	COLORREF OLECOLOR2COLORREF(OLE_COLOR oleColor, HPALETTE hPalette = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies all values to all associated controls</em>
	///
	/// \sa LoadSettings
	void ApplySettings(void);
	/// \brief <em>Stores the controls' contents as the specified panel's properties</em>
	///
	/// \param[in] panelIndex The panel to save.
	///
	/// \sa LoadPanelSettings
	void ApplyPanelSettings(int panelIndex);
	/// \brief <em>Loads the current properties' settings from the associated controls</em>
	///
	/// \sa ApplySettings
	void LoadSettings(void);
	/// \brief <em>Fills the controls with the specified panel's properties</em>
	///
	/// \param[in] panelIndex The panel to load.
	///
	/// \sa ApplyPanelSettings
	void LoadPanelSettings(int panelIndex);

	/// \brief <em>Holds wrappers for the dialog's GUI elements</em>
	struct Controls
	{
		/// \brief <em>Wraps the edit box displaying the selected panel's index</em>
		CEdit panelIndexEdit;
		/// \brief <em>Wraps the up-down button used to select the panels</em>
		CUpDownCtrl panelIndexUpDown;
		/// \brief <em>Wraps the button used to insert a panel</em>
		CButton insertPanelButton;
		/// \brief <em>Wraps the button used to remove the selected panel</em>
		CButton removePanelButton;
		/// \brief <em>Wraps the edit box displaying the selected panel's \c Text property</em>
		CEdit panelTextEdit;
		/// \brief <em>Wraps the edit box displaying the selected panel's \c ToolTipText property</em>
		CEdit panelToolTipTextEdit;
		/// \brief <em>Wraps the combo box displaying the selected panel's \c BorderStyle property</em>
		CComboBox panelBorderStyleCombo;
		/// \brief <em>Wraps the edit box displaying the selected panel's \c MinimumWidth property</em>
		CEdit panelMinWidthEdit;
		/// \brief <em>Wraps the combo box displaying the selected panel's \c Alignment property</em>
		CComboBox panelAlignmentCombo;
		/// \brief <em>Wraps the edit box displaying the selected panel's \c PreferredWidth property</em>
		CEdit panelPrefWidthEdit;
		/// \brief <em>Wraps the combo box displaying the selected panel's \c Content property</em>
		CComboBox panelContentCombo;
		/// \brief <em>Wraps the edit box displaying the selected panel's \c PanelData property</em>
		CEdit panelDataEdit;
		/// \brief <em>Wraps the radio button used to set the selected panel's \c ForeColor property to -1</em>
		CButton panelForeClrRadioDefault;
		/// \brief <em>Wraps the radio button used to set the selected panel's \c ForeColor property to a system color</em>
		CButton panelForeClrRadioSysClr;
		/// \brief <em>Wraps the radio button used to set the selected panel's \c ForeColor property to a RGB color</em>
		CButton panelForeClrRadioRGBClr;
		/// \brief <em>Wraps the combo box used to set the selected panel's \c ForeColor property to a system color</em>
		CComboBox panelForeColorCombo;
		/// \brief <em>Wraps the button used to set the selected panel's \c ForeColor property to a RGB color</em>
		///
		/// \sa ColorButton
		ColorButton panelForeColorButton;
		/// \brief <em>Wraps the check button used to make the selected panel the auto-sized panel</em>
		CButton panelAutoSizeCheck;
		/// \brief <em>Wraps the check button used to set the selected panel's \c Enabled property</em>
		CButton panelEnabledCheck;
		/// \brief <em>Wraps the check button used to set the selected panel's \c ParseTabs property</em>
		CButton panelParseTabsCheck;
		/// \brief <em>Wraps the check button used to set the selected panel's \c RightToLeftText property</em>
		CButton panelRTLTextCheck;
		/// \brief <em>Wraps the toolbar control providing the buttons to save/load settings</em>
		CToolBarCtrl toolbar;
		/// \brief <em>Holds the \c toolbar control's enabled icons</em>
		CImageList imagelistEnabled;
		/// \brief <em>Holds the \c toolbar control's disabled icons</em>
		CImageList imagelistDisabled;

		~Controls()
		{
			if(!imagelistEnabled.IsNull()) {
				imagelistEnabled.Destroy();
			}
			if(!imagelistDisabled.IsNull()) {
				imagelistDisabled.Destroy();
			}
		}
	} controls;

	/// \brief <em>Holds the dialog's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the \c panelTextEdit control's content was edited</em>
		UINT changedPanelText : 1;
		/// \brief <em>If \c TRUE, the \c panelToolTipTextEdit control's content was edited</em>
		UINT changedPanelToolTipText : 1;
		/// \brief <em>If \c TRUE, the \c panelBorderStyleCombo control's content was edited</em>
		UINT changedPanelBorderStyle : 1;
		/// \brief <em>If \c TRUE, the \c panelMinWidthEdit control's content was edited</em>
		UINT changedPanelMinimumWidth : 1;
		/// \brief <em>If \c TRUE, the \c panelAlignmentCombo control's content was edited</em>
		UINT changedPanelAlignment : 1;
		/// \brief <em>If \c TRUE, the \c panelPrefWidthEdit control's content was edited</em>
		UINT changedPanelPreferredWidth : 1;
		/// \brief <em>If \c TRUE, the \c panelContentCombo control's content was edited</em>
		UINT changedPanelContent : 1;
		/// \brief <em>If \c TRUE, the \c panelDataEdit control's content was edited</em>
		UINT changedPanelData : 1;
		/// \brief <em>If \c TRUE, the \c ForeColor settings were edited</em>
		UINT changedPanelForeColor : 1;
		/// \brief <em>If \c TRUE, the \c panelAutoSizeCheck control was clicked</em>
		UINT changedPanelToAutoSize : 1;
		/// \brief <em>If \c TRUE, the \c panelEnabledCheck control was clicked</em>
		UINT changedPanelEnabled : 1;
		/// \brief <em>If \c TRUE, the \c panelParseTabsCheck control was clicked</em>
		UINT changedPanelParseTabs : 1;
		/// \brief <em>If \c TRUE, the \c panelRTLTextCheck control was clicked</em>
		UINT changedPanelRTLText : 1;
		/// \brief <em>Holds the index of the panel whose properties are currently displayed</em>
		int currentPanelIndex;

		Flags()
		{
			changedPanelText = FALSE;
			changedPanelToolTipText = FALSE;
			changedPanelBorderStyle = FALSE;
			changedPanelMinimumWidth = FALSE;
			changedPanelAlignment = FALSE;
			changedPanelPreferredWidth = FALSE;
			changedPanelContent = FALSE;
			changedPanelData = FALSE;
			changedPanelForeColor = FALSE;
			changedPanelToAutoSize = FALSE;
			changedPanelEnabled = FALSE;
			changedPanelParseTabs = FALSE;
			changedPanelRTLText = FALSE;
			currentPanelIndex = -1;
		}
	} flags;

private:
};     // PanelProperties

OBJECT_ENTRY_AUTO(CLSID_PanelProperties, PanelProperties)