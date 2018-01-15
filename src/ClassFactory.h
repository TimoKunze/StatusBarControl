//////////////////////////////////////////////////////////////////////
/// \class ClassFactory
/// \author Timo "TimoSoft" Kunze
/// \brief <em>A helper class for creating special objects</em>
///
/// This class provides methods to create objects and initialize them with given values.
///
/// \todo Improve documentation.
///
/// \sa StatusBar
//////////////////////////////////////////////////////////////////////


#pragma once

#include "StatusBarPanel.h"
#include "StatusBarPanels.h"
#include "TargetOLEDataObject.h"


class ClassFactory
{
public:
	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's \c IOLEDataObject implementation as a smart pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	///
	/// \return A smart pointer to the created object's \c IOLEDataObject implementation.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static CComPtr<IOLEDataObject> InitOLEDataObject(IDataObject* pDataObject)
	{
		CComPtr<IOLEDataObject> pOLEDataObj = NULL;
		InitOLEDataObject(pDataObject, IID_IOLEDataObject, reinterpret_cast<LPUNKNOWN*>(&pOLEDataObj));
		return pOLEDataObj;
	};

	/// \brief <em>Creates an \c OLEDataObject object</em>
	///
	/// \overload
	///
	/// Creates an \c OLEDataObject object that wraps an object implementing \c IDataObject and returns
	/// the created object's implementation of a given interface as a raw pointer.
	///
	/// \param[in] pDataObject The \c IDataObject implementation the \c OLEDataObject object will work
	///            on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppDataObject Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>
	static void InitOLEDataObject(IDataObject* pDataObject, REFIID requiredInterface, LPUNKNOWN* ppDataObject)
	{
		ATLASSERT_POINTER(ppDataObject, LPUNKNOWN);
		ATLASSUME(ppDataObject);

		*ppDataObject = NULL;

		// create an OLEDataObject object and initialize it with the given parameters
		CComObject<TargetOLEDataObject>* pOLEDataObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObj);
		pOLEDataObj->AddRef();
		pOLEDataObj->Attach(pDataObject);
		pOLEDataObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppDataObject));
		pOLEDataObj->Release();
	};

	/// \brief <em>Creates a \c StatusBarPanel object</em>
	///
	/// Creates a \c StatusBarPanel object that represents a given status bar panel and returns
	/// its \c IStatusBarPanel implementation as a smart pointer.
	///
	/// \param[in] panelIndex The panel for which to create the object.
	/// \param[in] pStatBar The \c StatusBar object the \c StatusBarPanel object will work on.
	/// \param[in] validatePanelIndex If \c TRUE, the method will check whether the \c panelIndex
	///            parameter identifies an existing panel; otherwise not.
	///
	/// \return A smart pointer to the created object's \c IStatusBarPanel implementation.
	static inline CComPtr<IStatusBarPanel> InitPanel(int panelIndex, StatusBar* pStatBar, BOOL validatePanelIndex = TRUE)
	{
		CComPtr<IStatusBarPanel> pPanel = NULL;
		InitPanel(panelIndex, pStatBar, IID_IStatusBarPanel, reinterpret_cast<LPUNKNOWN*>(&pPanel), validatePanelIndex);
		return pPanel;
	};

	/// \brief <em>Creates a \c StatusBarPanel object</em>
	///
	/// \overload
	///
	/// Creates a \c StatusBarPanel object that represents a given status bar panel and returns
	/// its implementation of a given interface as a raw pointer.
	///
	/// \param[in] panelIndex The panel for which to create the object.
	/// \param[in] pStatBar The \c StatusBar object the \c StatusBarPanel object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppPanel Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	/// \param[in] validatePanelIndex If \c TRUE, the method will check whether the \c panelIndex
	///            parameter identifies an existing panel; otherwise not.
	static inline void InitPanel(int panelIndex, StatusBar* pStatBar, REFIID requiredInterface, LPUNKNOWN* ppPanel, BOOL validatePanelIndex = TRUE)
	{
		ATLASSERT_POINTER(ppPanel, LPUNKNOWN);
		ATLASSUME(ppPanel);

		*ppPanel = NULL;
		if(validatePanelIndex && !IsValidPanelIndex(panelIndex, pStatBar)) {
			// there's nothing useful we could return
			return;
		}

		// create a StatusBarPanel object and initialize it with the given parameters
		CComObject<StatusBarPanel>* pStatBarPanelObj = NULL;
		CComObject<StatusBarPanel>::CreateInstance(&pStatBarPanelObj);
		pStatBarPanelObj->AddRef();
		pStatBarPanelObj->SetOwner(pStatBar);
		pStatBarPanelObj->Attach(panelIndex);
		pStatBarPanelObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppPanel));
		pStatBarPanelObj->Release();
	};

	/// \brief <em>Creates a \c StatusBarPanels object</em>
	///
	/// Creates a \c StatusBarPanels object that represents a collection of status bar panels and
	/// returns its \c IStatusBarPanels implementation as a smart pointer.
	///
	/// \param[in] pStatBar The \c StatusBar object the \c StatusBarPanels object will work on.
	///
	/// \return A smart pointer to the created object's \c IStatusBarPanels implementation.
	static inline CComPtr<IStatusBarPanels> InitPanels(StatusBar* pStatBar)
	{
		CComPtr<IStatusBarPanels> pPanels = NULL;
		InitPanels(pStatBar, IID_IStatusBarPanels, reinterpret_cast<LPUNKNOWN*>(&pPanels));
		return pPanels;
	};

	/// \brief <em>Creates a \c StatusBarPanels object</em>
	///
	/// \overload
	///
	/// Creates a \c StatusBarPanels object that represents a collection of status bar panels and
	/// returns its implementation of a given interface as a raw pointer.
	///
	/// \param[in] pStatBar The \c StatusBar object the \c StatusBarPanels object will work on.
	/// \param[in] requiredInterface The IID of the interface of which the object's implementation will
	///            be returned.
	/// \param[out] ppPanels Receives the object's implementation of the interface identified by
	///             \c requiredInterface.
	static inline void InitPanels(StatusBar* pStatBar, REFIID requiredInterface, LPUNKNOWN* ppPanels)
	{
		ATLASSERT_POINTER(ppPanels, LPUNKNOWN);
		ATLASSUME(ppPanels);

		*ppPanels = NULL;
		// create a StatusBarPanels object and initialize it with the given parameters
		CComObject<StatusBarPanels>* pStatBarPanelsObj = NULL;
		CComObject<StatusBarPanels>::CreateInstance(&pStatBarPanelsObj);
		pStatBarPanelsObj->AddRef();
		pStatBarPanelsObj->SetOwner(pStatBar);
		pStatBarPanelsObj->QueryInterface(requiredInterface, reinterpret_cast<LPVOID*>(ppPanels));
		pStatBarPanelsObj->Release();
	};
};     // ClassFactory