//////////////////////////////////////////////////////////////////////
/// \class StatusBarPanel
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing status bar panel</em>
///
/// This class is a wrapper around a status bar panel within the control.
///
/// \if UNICODE
///   \sa StatBarLibU::IStatusBarPanel, StatusBarPanels, StatusBar
/// \else
///   \sa StatBarLibA::IStatusBarPanel, StatusBarPanels, StatusBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "StatBarU.h"
#else
	#include "StatBarA.h"
#endif
#include "_IStatusBarPanelEvents_CP.h"
#include "helpers.h"
#include "StatusBar.h"


class ATL_NO_VTABLE StatusBarPanel : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<StatusBarPanel, &CLSID_StatusBarPanel>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<StatusBarPanel>,
    public Proxy_IStatusBarPanelEvents<StatusBarPanel>, 
    #ifdef UNICODE
    	public IDispatchImpl<IStatusBarPanel, &IID_IStatusBarPanel, &LIBID_StatBarLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IStatusBarPanel, &IID_IStatusBarPanel, &LIBID_StatBarLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<StatusBarPanel>,
    public IPersistStorageImpl<StatusBarPanel>,
    public IPersistPropertyBagImpl<StatusBarPanel>
{
	friend class ClassFactory;
	friend class StatusBar;
	friend class StatusBarPanels;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		// required for persistence
		BOOL m_bRequiresSave;

		DECLARE_REGISTRY_RESOURCEID(IDR_STATUSBARPANEL)

		BEGIN_COM_MAP(StatusBarPanel)
			COM_INTERFACE_ENTRY(IStatusBarPanel)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IPersistStorage)
		END_COM_MAP()

		BEGIN_PROP_MAP(StatusBarPanel)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(StatusBarPanel)
			CONNECTION_POINT_ENTRY(__uuidof(_IStatusBarPanelEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportsErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to implement conditional persistence</em>
	///
	/// We want to implement conditional persistence. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the object's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to implement conditional persistence</em>
	///
	/// We want to implement conditional persistence. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the object's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the object should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the object should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to implement conditional persistence</em>
	///
	/// We want to implement conditional persistence. This can't be done by just using ATL's persistence
	/// macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the object's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to implement conditional persistence</em>
	///
	/// We want to implement conditional persistence. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistStreamInitImpl::Load and read directly from the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the object's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to implement conditional persistence</em>
	///
	/// We want to implement conditional persistence. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistStreamInitImpl::Save and write directly into the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the object's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the object should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IStatusBarPanel
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c Alignment property</em>
	///
	/// Retrieves the alignment of the panel's text. Any of the values defined by the \c AlignmentConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Alignment, get_Text, get_ParseTabs, StatBarLibU::AlignmentConstants
	/// \else
	///   \sa put_Alignment, get_Text, get_ParseTabs, StatBarLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Alignment(AlignmentConstants* pValue);
	/// \brief <em>Sets the \c Alignment property</em>
	///
	/// Sets the alignment of the panel's text. Any of the values defined by the \c AlignmentConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Alignment, put_Text, put_ParseTabs, StatBarLibU::AlignmentConstants
	/// \else
	///   \sa get_Alignment, put_Text, put_ParseTabs, StatBarLibA::AlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Alignment(AlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of border that is drawn around the panel. Any of the values defined by the
	/// \c PanelBorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, StatusBar::get_BorderStyle, StatBarLibU::PanelBorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, StatusBar::get_BorderStyle, StatBarLibA::PanelBorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(PanelBorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of border that is drawn around the panel. Any of the values defined by the
	/// \c PanelBorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, StatusBar::put_BorderStyle, StatBarLibU::PanelBorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, StatusBar::put_BorderStyle, StatBarLibA::PanelBorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(PanelBorderStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Content property</em>
	///
	/// Retrieves the panel's content type. Any of the values defined by the \c PanelContentConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \if UNICODE
	///   \sa put_Content, StatBarLibU::PanelContentConstants
	/// \else
	///   \sa put_Content, StatBarLibA::PanelContentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Content(PanelContentConstants* pValue);
	/// \brief <em>Sets the \c Content property</em>
	///
	/// Sets the panel's content type. Any of the values defined by the \c PanelContentConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \if UNICODE
	///   \sa get_Content, StatBarLibU::PanelContentConstants
	/// \else
	///   \sa get_Content, StatBarLibA::PanelContentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Content(PanelContentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentWidth property</em>
	///
	/// Retrieves the panel's current width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa StatusBar::get_PanelToAutoSize, get_PreferredWidth, get_MinimumWidth
	virtual HRESULT STDMETHODCALLTYPE get_CurrentWidth(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether this panel's content is drawn disabled. If set to \c VARIANT_TRUE, the content
	/// is drawn normal; otherwise it is drawn in disabled style.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa put_Enabled, get_hIcon, get_Text, get_Content, StatusBar::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Sets whether this panel's content is drawn disabled. If set to \c VARIANT_TRUE, the content
	/// is drawn normal; otherwise it is drawn in disabled style.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa get_Enabled, put_hIcon, put_Text, put_Content, StatusBar::put_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the panel's text color. If set to -1, the system's default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa put_ForeColor, StatusBar::get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the panel's text color. If set to -1, the system's default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa get_ForeColor, StatusBar::put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c hIcon property</em>
	///
	/// Retrieves the panel's icon handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hIcon, get_Text, get_Content
	virtual HRESULT STDMETHODCALLTYPE get_hIcon(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hIcon property</em>
	///
	/// Sets the panel's icon handle.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set icon does NOT get destroyed automatically.
	///
	/// \sa get_hIcon, put_Text, put_Content
	virtual HRESULT STDMETHODCALLTYPE put_hIcon(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves a zero-based index identifying this panel.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Adding or removing panels changes other panels' indices.\n
	///          This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumWidth property</em>
	///
	/// Retrieves the panel's minimum width in pixels. If the panel is automatically sized, its
	/// width won't decrease below the minimum width.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa put_MinimumWidth, StatusBar::get_PanelToAutoSize, get_CurrentWidth, get_PreferredWidth
	virtual HRESULT STDMETHODCALLTYPE get_MinimumWidth(LONG* pValue);
	/// \brief <em>Sets the \c MinimumWidth property</em>
	///
	/// Retrieves the panel's minimum width in pixels. If the panel is automatically sized, its
	/// width won't decrease below the minimum width.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa get_MinimumWidth, StatusBar::putref_PanelToAutoSize, put_PreferredWidth
	virtual HRESULT STDMETHODCALLTYPE put_MinimumWidth(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c PanelData property</em>
	///
	/// Retrieves the \c LONG value associated with the panel. Use this property to associate any data
	/// with the panel.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_PanelData
	virtual HRESULT STDMETHODCALLTYPE get_PanelData(LONG* pValue);
	/// \brief <em>Sets the \c PanelData property</em>
	///
	/// Sets the \c LONG value associated with the panel. Use this property to associate any data
	/// with the panel.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_PanelData
	virtual HRESULT STDMETHODCALLTYPE put_PanelData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ParseTabs property</em>
	///
	/// Retrieves whether the tabulators in this panel's text are parsed to create a multi-column panel.
	/// If set to \c VARIANT_TRUE, the tabulators are parsed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.
	///
	/// \sa put_ParseTabs, get_Text, get_Alignment
	virtual HRESULT STDMETHODCALLTYPE get_ParseTabs(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ParseTabs property</em>
	///
	/// Sets whether the tabulators in this panel's text are parsed to create a multi-column panel.
	/// If set to \c VARIANT_TRUE, the tabulators are parsed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 5.80 or higher.
	///
	/// \sa get_ParseTabs, put_Text, put_Alignment
	virtual HRESULT STDMETHODCALLTYPE put_ParseTabs(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c PreferredWidth property</em>
	///
	/// Retrieves the panel's preferred width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa put_PreferredWidth, StatusBar::get_PanelToAutoSize, get_CurrentWidth, get_MinimumWidth
	virtual HRESULT STDMETHODCALLTYPE get_PreferredWidth(LONG* pValue);
	/// \brief <em>Sets the \c PreferredWidth property</em>
	///
	/// Retrieves the panel's preferred width in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is not supported in simple mode.
	///
	/// \sa get_PreferredWidth, StatusBar::putref_PanelToAutoSize, put_MinimumWidth
	virtual HRESULT STDMETHODCALLTYPE put_PreferredWidth(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeftText property</em>
	///
	/// Retrieves the panel's reading direction. If set to \c VARIANT_TRUE, the panel's text reads from
	/// right to left; otherwise from left to right.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_RightToLeftText, StatusBar::get_RightToLeftLayout
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeftText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RightToLeftText property</em>
	///
	/// Sets the panel's reading direction. If set to \c VARIANT_TRUE, the panel's text reads from
	/// right to left; otherwise from left to right.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_RightToLeftText, StatusBar::put_RightToLeftLayout
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeftText(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the panel's text. Up to comctl32.dll version 6.0 the maximum number of characters in this
	/// text is specified by \c MAX_PANELTEXTLENGTH. Beginning with comctl32.dll version 6.10 the text length
	/// isn't limited anymore.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, MAX_PANELTEXTLENGTH, get_hIcon, get_Content, get_ParseTabs
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the panel's text. Up to comctl32.dll version 6.0 the maximum number of characters in this
	/// text is specified by \c MAX_PANELTEXTLENGTH. Beginning with comctl32.dll version 6.10 the text length
	/// isn't limited anymore.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, MAX_PANELTEXTLENGTH, put_hIcon, put_Content, put_ParseTabs
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c ToolTipText property</em>
	///
	/// Retrieves the panel's tooltip text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ToolTipText, StatusBar::get_ShowToolTips, StatusBar::Raise_SetupToolTipWindow
	virtual HRESULT STDMETHODCALLTYPE get_ToolTipText(BSTR* pValue);
	/// \brief <em>Sets the \c ToolTipText property</em>
	///
	/// Sets the panel's tooltip text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ToolTipText, StatusBar::put_ShowToolTips, StatusBar::Raise_SetupToolTipWindow
	virtual HRESULT STDMETHODCALLTYPE put_ToolTipText(BSTR newValue);

	/// \brief <em>Retrieves the panel's bounding rectangle</em>
	///
	/// Retrieves the panel's bounding rectangle (in pixels) within the control's client area.
	///
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the control's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the control's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the control's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method is not supported in simple mode.
	///
	/// \sa get_PreferredWidth
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given panel</em>
	///
	/// Attaches this object to a given panel, so that the panel's properties can be
	/// retrieved and set using this object's methods.
	///
	/// \param[in] panelIndex The panel to attach to.
	///
	/// \sa Detach
	void Attach(int panelIndex);
	/// \brief <em>Detaches this object from a panel</em>
	///
	/// Detaches this object from the panel it currently wraps, so that it doesn't wrap any
	/// panel anymore.
	///
	/// \sa Attach
	void Detach(void);

	/// \brief <em>Applies this object's settings to the specified panel</em>
	///
	/// \param[in] panelIndex The panel to which the settings will be applied.
	/// \param[in] textAndStyleOnly If set to \c TRUE, the method will apply the panel's text and styles
	///            only; otherwise it will also apply the tooltip text and the icon.
	///
	/// \sa ApplySettings, StatusBarPanels::Add, StatusBarPanels::Remove
	void ApplySettingsTo(int panelIndex, BOOL textAndStyleOnly);
	/// \brief <em>Applies this object's settings to its underlying panel</em>
	///
	/// \param[in] textAndStyleOnly If set to \c TRUE, the method will apply the panel's text and styles
	///            only; otherwise it will also apply the tooltip text and the icon.
	/// \param[in] updatePredefinedPanels If set to \c TRUE, the method will update the panel's text if
	///            it is a predefined one.
	///
	/// \sa ApplySettingsTo, StatusBarPanels::Add, StatusBarPanels::Remove
	void ApplySettings(BOOL textAndStyleOnly, BOOL updatePredefinedPanels);
	/// \brief <em>Sets the owner of this panel</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerStatBar
	void SetOwner(__in_opt StatusBar* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c StatusBar object that owns this panel</em>
		///
		/// \sa SetOwner
		StatusBar* pOwnerStatBar;
		/// \brief <em>The index of the panel wrapped by this object</em>
		int panelIndex;

		/// \brief <em>Holds the \c Alignment property's setting</em>
		///
		/// \sa get_Alignment, put_Alignment
		AlignmentConstants alignment;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		PanelBorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c Content property's setting</em>
		///
		/// \sa get_Content, put_Content
		PanelContentConstants content;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the \c hIcon property's setting</em>
		///
		/// \sa get_hIcon, put_hIcon
		HICON hIcon;
		/// \brief <em>Holds the \c MinimumWidth property's setting</em>
		///
		/// \sa get_MinimumWidth, put_MinimumWidth
		long minimumWidth;
		/// \brief <em>Holds the \c PanelData property's setting</em>
		///
		/// \sa get_PanelData, put_PanelData
		long panelData;
		/// \brief <em>Holds the \c ParseTabs property's setting</em>
		///
		/// \sa get_ParseTabs, put_ParseTabs
		UINT parseTabs : 1;
		/// \brief <em>Holds the \c PreferredWidth property's setting</em>
		///
		/// \sa get_PreferredWidth, put_PreferredWidth
		long preferredWidth;
		/// \brief <em>Holds the \c RightToLeftText property's setting</em>
		///
		/// \sa get_RightToLeftText, put_RightToLeftText
		UINT rightToLeftText : 1;
		/// \brief <em>Holds the \c Text property's setting</em>
		///
		/// \sa get_Text, put_Text
		CComBSTR text;
		/// \brief <em>Holds the \c ToolTipText property's setting</em>
		///
		/// \sa get_ToolTipText, put_ToolTipText
		CComBSTR toolTipText;

		Properties()
		{
			pOwnerStatBar = NULL;
			panelIndex = -1;

			alignment = alLeft;
			borderStyle = pbsNone;
			content = pcText;
			enabled = TRUE;
			foreColor = CLR_NONE;
			hIcon = NULL;
			minimumWidth = 100;
			panelData = 0;
			parseTabs = TRUE;
			preferredWidth = 100;
			rightToLeftText = FALSE;
			text = L"";
			toolTipText = L"";
		}

		~Properties();

		/// \brief <em>Retrieves the owning status bar's window handle</em>
		///
		/// \return The window handle of the status bar that contains this panel.
		///
		/// \sa pOwnerStatBar
		HWND GetStatBarHWnd(void);
	} properties;
};     // StatusBarPanel

OBJECT_ENTRY_AUTO(__uuidof(StatusBarPanel), StatusBarPanel)