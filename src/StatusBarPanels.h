//////////////////////////////////////////////////////////////////////
/// \class StatusBarPanels
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c StatusBarPanel objects</em>
///
/// This class provides easy access (including filtering) to collections of \c StatusBarPanel
/// objects. A \c StatusBarPanels object is used to group panels that have certain properties
/// in common.
///
/// \if UNICODE
///   \sa StatBarLibU::IStatusBarPanels, StatusBarPanel, StatusBar
/// \else
///   \sa StatBarLibA::IStatusBarPanels, StatusBarPanel, StatusBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "StatBarU.h"
#else
	#include "StatBarA.h"
#endif
#include "_IStatusBarPanelsEvents_CP.h"
#include "helpers.h"
#include "StatusBar.h"
#include "StatusBarPanel.h"


class ATL_NO_VTABLE StatusBarPanels : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<StatusBarPanels, &CLSID_StatusBarPanels>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<StatusBarPanels>,
    public Proxy_IStatusBarPanelsEvents<StatusBarPanels>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IStatusBarPanels, &IID_IStatusBarPanels, &LIBID_StatBarLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IStatusBarPanels, &IID_IStatusBarPanels, &LIBID_StatBarLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<StatusBarPanels>,
    public IPersistStorageImpl<StatusBarPanels>,
    public IPersistPropertyBagImpl<StatusBarPanels>
{
	friend class ClassFactory;
	friend class StatusBar;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		// required for persistence
		BOOL m_bRequiresSave;

		DECLARE_REGISTRY_RESOURCEID(IDR_STATUSBARPANELS)

		BEGIN_COM_MAP(StatusBarPanels)
			COM_INTERFACE_ENTRY(IStatusBarPanels)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_PROP_MAP(StatusBarPanels)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(StatusBarPanels)
			CONNECTION_POINT_ENTRY(__uuidof(_IStatusBarPanelsEvents))
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
	/// We want to persist some read-only properties. This can't be done by just using ATL's persistence
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
	/// We want to persist some read-only properties. This can't be done by just using ATL's persistence
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
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
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
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
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
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the panels</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c StatusBarPanel objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x panels</em>
	///
	/// Retrieves the next \c numberOfMaxItems panels from the iterator.
	///
	/// \param[in] numberOfMaxItems The maximum number of panels the array identified by \c pItems
	///            can contain.
	/// \param[in,out] pItems An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a panel's \c IStatusBarPanel implementation.
	/// \param[out] pNumberOfItemsReturned The number of panels that actually were copied to the
	///             array identified by \c pItems.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, StatusBarPanel,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// panel in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x panels</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfItemsToSkip panels.
	///
	/// \param[in] numberOfItemsToSkip The number of panels to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfItemsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IStatusBarPanels
	///
	//@{
	/// \brief <em>Retrieves a \c StatusBarPanel object from the collection</em>
	///
	/// Retrieves a \c StatusBarPanel object from the collection that wraps the status bar panel
	/// identified by \c panelIndex.
	///
	/// \param[in] panelIndex A value that identifies the status bar panel to be retrieved.
	/// \param[out] ppPanel Receives the panel's \c IStatusBarPanel implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa StatusBarPanel, Add, Remove
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG panelIndex, IStatusBarPanel** ppPanel);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c StatusBarPanel objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds a panel to the status bar</em>
	///
	/// Adds a panel with the specified properties at the specified position in the control and
	/// returns a \c StatusBarPanel object wrapping the inserted panel.
	///
	/// \param[in] panelText The new panel's text. Up to comctl32.dll version 6.0 the maximum number of
	///            characters in this text is specified by \c MAX_PANELTEXTLENGTH. Beginning with
	///            comctl32.dll version 6.10 the text length isn't limited anymore.
	/// \param[in] preferredWidth The new panel's preferred width in pixels.
	/// \param[in] content The new panel's content type. Any of the values defined by the
	///            \c PanelContentConstants enumeration is valid.
	/// \param[in] alignment The alignment of the new panel's text. Any of the values defined by the
	///            \c AlignmentConstants enumeration is valid.
	/// \param[in] borderStyle The kind of border that is drawn around the new panel. Any of the values
	///            defined by the \c PanelBorderStyleConstants enumeration is valid.
	/// \param[in] panelData A \c Long value that will be associated with the panel.
	/// \param[in] insertAt The new panel's zero-based index. If set to -1, the panel will be inserted
	///            as the last panel.
	/// \param[out] ppAddedPanel Receives the added panel's \c IStatusBarPanel implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, StatusBarPanel::get_Text, StatusBarPanel::get_PreferredWidth,
	///       StatusBarPanel::get_Content, StatusBarPanel::get_Alignment, StatusBarPanel::get_BorderStyle,
	///       StatusBarPanel::get_PanelData, StatusBarPanel::get_Index, StatusBar::get_PanelToAutoSize,
	///       MAX_PANELTEXTLENGTH, StatBarLibU::PanelContentConstants, StatBarLibU::AlignmentConstants,
	///       StatBarLibU::PanelBorderStyleConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, StatusBarPanel::get_Text, StatusBarPanel::get_PreferredWidth,
	///       StatusBarPanel::get_Content, StatusBarPanel::get_Alignment, StatusBarPanel::get_BorderStyle,
	///       StatusBarPanel::get_PanelData, StatusBarPanel::get_Index, StatusBar::get_PanelToAutoSize,
	///       MAX_PANELTEXTLENGTH, StatBarLibA::PanelContentConstants, StatBarLibA::AlignmentConstants,
	///       StatBarLibA::PanelBorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(BSTR panelText, LONG preferredWidth, PanelContentConstants content = pcText, AlignmentConstants alignment = alLeft, PanelBorderStyleConstants borderStyle = pbsSunken, LONG panelData = 0, LONG insertAt = -1, IStatusBarPanel** ppAddedPanel = NULL);
	/// \brief <em>Counts the panels in the collection</em>
	///
	/// Retrieves the number of \c StatusBarPanel objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified panel in the collection from the status bar</em>
	///
	/// \param[in] panelIndex A value that identifies the status bar panel to be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The control always contains at least one panel.
	///
	/// \sa Add, Count, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG panelIndex);
	/// \brief <em>Removes all panels in the collection from the status bar</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The control always contains at least one panel.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies each contained \c StatusBarPanel object's settings to its underlying panel</em>
	///
	/// \param[in] textAndStyleOnly If set to \c TRUE, the method will apply the panels' texts and styles
	///            only; otherwise it will also apply the tooltip texts and the icons.
	/// \param[in] updatePredefinedPanels If set to \c TRUE, the method will update predefined panels' text.
	///
	/// \sa StatusBar::ApplyPanelWidths
	void ApplySettings(BOOL textAndStyleOnly, BOOL updatePredefinedPanels);

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerStatBar
	void SetOwner(__in_opt StatusBar* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c StatusBar object that owns this collection</em>
		///
		/// \sa SetOwner
		StatusBar* pOwnerStatBar;
		#ifdef USE_STL
			/// \brief <em>Points to the next enumerated panel</em>
			std::list< CComObject<StatusBarPanel>* >::iterator nextEnumeratedPanel;
			/// \brief <em>A list containing all \c StatusBarPanel objects in this collection</em>
			std::list< CComObject<StatusBarPanel>* > panels;
		#else
			/// \brief <em>Points to the next enumerated panel</em>
			POSITION nextEnumeratedPanel;
			/// \brief <em>A list containing all \c StatusBarPanel objects in this collection</em>
			CAtlList< CComObject<StatusBarPanel>* > panels;
		#endif

		Properties()
		{
			pOwnerStatBar = NULL;
			#ifdef USE_STL
				nextEnumeratedPanel = panels.begin();
			#else
				nextEnumeratedPanel = panels.GetHeadPosition();
			#endif
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another struct of the same type</em>
		///
		/// \param[in] target The target struct that the content will be copied to.
		///
		/// \sa Clone
		void CopyTo(Properties& target);

		/// \brief <em>Retrieves the owning status bar's window handle</em>
		///
		/// \return The window handle of the status bar that contains the panels in this collection.
		///
		/// \sa pOwnerStatBar
		HWND GetStatBarHWnd(void);
	} properties;
};     // StatusBarPanels

OBJECT_ENTRY_AUTO(__uuidof(StatusBarPanels), StatusBarPanels)