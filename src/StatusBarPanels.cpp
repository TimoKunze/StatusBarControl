// StatusBarPanels.cpp: Manages a collection of StatusBarPanel objects

#include "stdafx.h"
#include "StatusBarPanels.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportsErrorInfo
STDMETHODIMP StatusBarPanels::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IStatusBarPanels, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportsErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP StatusBarPanels::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	ATLASSUME(pPropertyBag);
	if(!pPropertyBag) {
		return E_POINTER;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_I4;
	if(SUCCEEDED(pPropertyBag->Read(OLESTR("Version"), &propertyValue, pErrorLog))) {
		if(propertyValue.lVal > 0x0101) {
			return E_FAIL;
		}
	}

	// read the number of panels
	int numberOfPanels = 0;
	HRESULT hr = pPropertyBag->Read(OLESTR("NumPanels"), &propertyValue, pErrorLog);
	if(FAILED(hr)) {
		return hr;
	}
	numberOfPanels = propertyValue.lVal;

	// release the current objects
	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::const_iterator it = properties.panels.begin();
		while(it != properties.panels.end()) {
			(*it)->Release();
			++it;
		}
		properties.panels.clear();
	#else
		POSITION p = properties.panels.GetHeadPosition();
		while(p) {
			properties.panels.GetAt(p)->Release();
			properties.panels.GetNext(p);
		}
		properties.panels.RemoveAll();
	#endif

	// read the panel objects
	OLECHAR entryName[50];
	for(int panelIndex = 1; panelIndex <= numberOfPanels; ++panelIndex) {
		propertyValue.vt = VT_DISPATCH;
		propertyValue.pdispVal = NULL;

		CComObject<StatusBarPanel>* pStatBarPanelObj = NULL;
		CComObject<StatusBarPanel>::CreateInstance(&pStatBarPanelObj);
		pStatBarPanelObj->AddRef();
		pStatBarPanelObj->SetOwner(properties.pOwnerStatBar);
		pStatBarPanelObj->Attach(panelIndex - 1);
		pStatBarPanelObj->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));

		ZeroMemory(entryName, 50 * sizeof(OLECHAR));
		ATLVERIFY(SUCCEEDED(StringCchPrintfW(entryName, 50, L"Panel%i", panelIndex)));

		hr = pPropertyBag->Read(entryName, &propertyValue, pErrorLog);
		if(FAILED(hr)) {
			return hr;
		}
		#ifdef USE_STL
			properties.panels.push_back(pStatBarPanelObj);
		#else
			properties.panels.AddTail(pStatBarPanelObj);
		#endif

		VariantClear(&propertyValue);
	}

	m_bRequiresSave = FALSE;
	return S_OK;
}

STDMETHODIMP StatusBarPanels::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL /*saveAllProperties*/)
{
	ATLASSUME(pPropertyBag);
	if(!pPropertyBag) {
		return E_POINTER;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	HRESULT hr = S_OK;
	propertyValue.vt = VT_I4;
	propertyValue.lVal = 0x0101;
	if(FAILED(hr = pPropertyBag->Write(OLESTR("Version"), &propertyValue))) {
		return hr;
	}

	// store the number of panels
	#ifdef USE_STL
		propertyValue.lVal = properties.panels.size();
	#else
		propertyValue.lVal = properties.panels.GetCount();
	#endif
	if(FAILED(hr = pPropertyBag->Write(OLESTR("NumPanels"), &propertyValue))) {
		return hr;
	}

	// store the panel objects
	OLECHAR entryName[50];
	#ifdef USE_STL
		for(std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin(); it != properties.panels.end(); ++it) {
			propertyValue.vt = VT_DISPATCH;
			propertyValue.pdispVal = NULL;
			(*it)->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));

			ZeroMemory(entryName, 50 * sizeof(OLECHAR));
			ATLVERIFY(SUCCEEDED(StringCchPrintfW(entryName, 50, L"Panel%i", std::distance(properties.panels.begin(), it) + 1)));
	#else
		POSITION p = properties.panels.GetHeadPosition();
		int i = 1;
		while(p) {
			propertyValue.vt = VT_DISPATCH;
			propertyValue.pdispVal = NULL;
			properties.panels.GetAt(p)->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&propertyValue.pdispVal));

			ZeroMemory(entryName, 50 * sizeof(OLECHAR));
			ATLVERIFY(SUCCEEDED(StringCchPrintfW(entryName, 50, L"Panel%i", i)));
	#endif

		if(FAILED(hr = pPropertyBag->Write(entryName, &propertyValue))) {
			VariantClear(&propertyValue);
			return hr;
		}
		VariantClear(&propertyValue);

		#ifndef USE_STL
			properties.panels.GetNext(p);
			++i;
		#endif
	}

	if(clearDirtyFlag) {
		m_bRequiresSave = FALSE;
	}
	return S_OK;
}

STDMETHODIMP StatusBarPanels::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*version*/) + sizeof(DWORD/*atlVersion*/);

	// we've 1 VT_I4 property...
	pSize->QuadPart += sizeof(VARTYPE) + sizeof(LONG);

	// ...and n VT_DISPATCH properties, each of which we've to query for its size
	#ifdef USE_STL
		for(std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin(); it != properties.panels.end(); ++it) {
	#else
		POSITION p = properties.panels.GetHeadPosition();
		while(p) {
	#endif
		// the marker
		pSize->QuadPart += sizeof(VARTYPE) + sizeof(CLSID);

		CComPtr<IPersistStreamInit> pStreamInit = NULL;
		#ifdef USE_STL
			if(FAILED((*it)->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
				(*it)->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
			}
		#else
			if(FAILED(properties.panels.GetAt(p)->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
				properties.panels.GetAt(p)->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
			}
		#endif
		if(pStreamInit) {
			ULARGE_INTEGER tmp = {0};
			pStreamInit->GetSizeMax(&tmp);
			pSize->QuadPart += tmp.QuadPart;
		}

		#ifndef USE_STL
			properties.panels.GetNext(p);
		#endif
	}

	return S_OK;
}

STDMETHODIMP StatusBarPanels::Load(LPSTREAM pStream)
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

	// read the number of panels
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	int numberOfPanels = propertyValue.lVal;

	// release the current objects
	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::const_iterator it = properties.panels.begin();
		while(it != properties.panels.end()) {
			(*it)->Release();
			++it;
		}
		properties.panels.clear();
	#else
		POSITION p = properties.panels.GetHeadPosition();
		while(p) {
			properties.panels.GetAt(p)->Release();
			properties.panels.GetNext(p);
		}
		properties.panels.RemoveAll();
	#endif

	// read the panel objects
	for(int panelIndex = 1; panelIndex <= numberOfPanels; ++panelIndex) {
		VARTYPE vt;
		if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
			return hr;
		}
		CLSID clsid;
		if(FAILED(hr = ReadClassStm(pStream, &clsid))) {
			return hr;
		}

		CComObject<StatusBarPanel>* pStatBarPanelObj = NULL;
		CComObject<StatusBarPanel>::CreateInstance(&pStatBarPanelObj);
		pStatBarPanelObj->AddRef();
		pStatBarPanelObj->SetOwner(properties.pOwnerStatBar);
		pStatBarPanelObj->Attach(panelIndex - 1);

		CComPtr<IPersistStreamInit> pStreamInit = NULL;
		if(FAILED(hr = pStatBarPanelObj->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			if(FAILED(hr = pStatBarPanelObj->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
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
		#ifdef USE_STL
			properties.panels.push_back(pStatBarPanelObj);
		#else
			properties.panels.AddTail(pStatBarPanelObj);
		#endif
	}

	m_bRequiresSave = FALSE;
	return S_OK;
}

STDMETHODIMP StatusBarPanels::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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

	// store the number of panels
	propertyValue.vt = VT_I4;
	#ifdef USE_STL
		propertyValue.lVal = properties.panels.size();
	#else
		propertyValue.lVal = properties.panels.GetCount();
	#endif
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	// store the panel objects
	#ifdef USE_STL
		for(std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin(); it != properties.panels.end(); ++it) {
			CComPtr<IPersistStreamInit> pStreamInit = NULL;
			if(FAILED(hr = (*it)->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
				if(FAILED(hr = (*it)->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
					return hr;
				}
			}
	#else
		POSITION p = properties.panels.GetHeadPosition();
		while(p) {
			CComPtr<IPersistStreamInit> pStreamInit = NULL;
			if(FAILED(hr = properties.panels.GetAt(p)->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
				if(FAILED(hr = properties.panels.GetAt(p)->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
					return hr;
				}
			}
	#endif

			// store some marker
			VARTYPE vt = VT_DISPATCH;
			if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
				return hr;
			}
			if(pStreamInit) {
				CLSID clsid = CLSID_NULL;
				if(FAILED(hr = pStreamInit->GetClassID(&clsid))) {
					return hr;
				}
				if(FAILED(hr = WriteClassStm(pStream, clsid))) {
					return hr;
				}
				if(FAILED(hr = pStreamInit->Save(pStream, clearDirtyFlag))) {
					return hr;
				}
			} else {
				if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
					return hr;
				}
			}

		#ifndef USE_STL
			properties.panels.GetNext(p);
		#endif
	}

	if(clearDirtyFlag) {
		m_bRequiresSave = FALSE;
	}
	return S_OK;
}


//////////////////////////////////////////////////////////////////////
// implementation of IEnumVARIANT
STDMETHODIMP StatusBarPanels::Clone(IEnumVARIANT** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPENUMVARIANT);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	*ppEnumerator = NULL;
	CComObject<StatusBarPanels>* pStatBarPanelsObj = NULL;
	CComObject<StatusBarPanels>::CreateInstance(&pStatBarPanelsObj);
	pStatBarPanelsObj->AddRef();

	// clone all settings
	properties.CopyTo(pStatBarPanelsObj->properties);

	pStatBarPanelsObj->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<LPVOID*>(ppEnumerator));
	pStatBarPanelsObj->Release();
	return S_OK;
}

STDMETHODIMP StatusBarPanels::Next(ULONG numberOfMaxItems, VARIANT* pItems, ULONG* pNumberOfItemsReturned)
{
	ATLASSERT_POINTER(pItems, VARIANT);
	if(!pItems) {
		return E_POINTER;
	}

	ULONG i = 0;
	for(i = 0; i < numberOfMaxItems; ++i) {
		VariantInit(&pItems[i]);

		#ifdef USE_STL
			if(std::distance(properties.nextEnumeratedPanel, properties.panels.end()) <= 0) {
				// there's nothing more to iterate
				break;
			}

			(*properties.nextEnumeratedPanel)->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
			pItems[i].vt = VT_DISPATCH;
			++properties.nextEnumeratedPanel;
		#else
			if(!properties.nextEnumeratedPanel) {
				// there's nothing more to iterate
				break;
			}

			properties.panels.GetAt(properties.nextEnumeratedPanel)->QueryInterface(IID_IDispatch, reinterpret_cast<LPVOID*>(&pItems[i].pdispVal));
			pItems[i].vt = VT_DISPATCH;
			properties.panels.GetNext(properties.nextEnumeratedPanel);
		#endif
	}
	if(pNumberOfItemsReturned) {
		*pNumberOfItemsReturned = i;
	}

	return (i == numberOfMaxItems ? S_OK : S_FALSE);
}

STDMETHODIMP StatusBarPanels::Reset(void)
{
	#ifdef USE_STL
		properties.nextEnumeratedPanel = properties.panels.begin();
	#else
		properties.nextEnumeratedPanel = properties.panels.GetHeadPosition();
	#endif
	return S_OK;
}

STDMETHODIMP StatusBarPanels::Skip(ULONG numberOfItemsToSkip)
{
	#ifdef USE_STL
		if(properties.nextEnumeratedPanel != properties.panels.end()) {
			std::advance(properties.nextEnumeratedPanel, numberOfItemsToSkip);
		}
	#else
		if(properties.nextEnumeratedPanel) {
			for(; numberOfItemsToSkip > 0; --numberOfItemsToSkip) {
				properties.panels.GetNext(properties.nextEnumeratedPanel);
			}
		}
	#endif
	return S_OK;
}
// implementation of IEnumVARIANT
//////////////////////////////////////////////////////////////////////


StatusBarPanels::Properties::~Properties()
{
	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::const_iterator it = panels.begin();
		while(it != panels.end()) {
			(*it)->Release();
			++it;
		}
		panels.clear();
	#else
		POSITION p = panels.GetHeadPosition();
		while(p) {
			panels.GetAt(p)->Release();
			panels.GetNext(p);
		}
		panels.RemoveAll();
	#endif
	if(pOwnerStatBar) {
		pOwnerStatBar->Release();
	}
}

void StatusBarPanels::Properties::CopyTo(Properties& target)
{
	target.pOwnerStatBar = pOwnerStatBar;
	if(target.pOwnerStatBar) {
		target.pOwnerStatBar->AddRef();
	}

	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::const_iterator it = panels.begin();
		while(it != panels.end()) {
			(*it)->AddRef();
			target.panels.push_back(*it);
			++it;
		}

		target.nextEnumeratedPanel = target.panels.begin();
		std::advance(target.nextEnumeratedPanel, std::distance(panels.begin(), nextEnumeratedPanel));
	#else
		POSITION p = panels.GetHeadPosition();
		while(p) {
			panels.GetAt(p)->AddRef();
			target.panels.AddTail(panels.GetAt(p));
			panels.GetNext(p);
		}

		if(nextEnumeratedPanel) {
			target.nextEnumeratedPanel = target.panels.GetHeadPosition();
			for(p = nextEnumeratedPanel; p != panels.GetHeadPosition(); panels.GetPrev(p), target.panels.GetNext(target.nextEnumeratedPanel));
		} else {
			target.nextEnumeratedPanel = NULL;
		}
	#endif
}

HWND StatusBarPanels::Properties::GetStatBarHWnd(void)
{
	ATLASSUME(pOwnerStatBar);

	OLE_HANDLE handle = NULL;
	pOwnerStatBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


void StatusBarPanels::SetOwner(StatusBar* pOwner)
{
	if(properties.pOwnerStatBar) {
		properties.pOwnerStatBar->Release();
	}
	properties.pOwnerStatBar = pOwner;
	if(properties.pOwnerStatBar) {
		properties.pOwnerStatBar->AddRef();
	}
}


void StatusBarPanels::ApplySettings(BOOL textAndStyleOnly, BOOL updatePredefinedPanels)
{
	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin();
		while(it != properties.panels.end()) {
			(*it)->ApplySettings(textAndStyleOnly, updatePredefinedPanels);
			++it;
		}
	#else
		POSITION p = properties.panels.GetHeadPosition();
		while(p) {
			properties.panels.GetAt(p)->ApplySettings(textAndStyleOnly, updatePredefinedPanels);
			properties.panels.GetNext(p);
		}
	#endif
}


STDMETHODIMP StatusBarPanels::get_Item(LONG panelIndex, IStatusBarPanel** ppPanel)
{
	ATLASSERT_POINTER(ppPanel, IStatusBarPanel*);
	if(!ppPanel) {
		return E_POINTER;
	}

	#ifdef USE_STL
		if((panelIndex < 0) || (panelIndex >= static_cast<LONG>(properties.panels.size()))) {
			return DISP_E_BADINDEX;
		}

		std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin();
		std::advance(it, panelIndex);
		(*it)->QueryInterface(IID_IStatusBarPanel, reinterpret_cast<LPVOID*>(ppPanel));
		return S_OK;
	#else
		POSITION p = properties.panels.FindIndex(panelIndex);
		if(p) {
			properties.panels.GetAt(p)->QueryInterface(IID_IStatusBarPanel, reinterpret_cast<LPVOID*>(ppPanel));
			return S_OK;
		}
		*ppPanel = NULL;
		return DISP_E_BADINDEX;
	#endif
}

STDMETHODIMP StatusBarPanels::get__NewEnum(IUnknown** ppEnumerator)
{
	ATLASSERT_POINTER(ppEnumerator, LPUNKNOWN);
	if(!ppEnumerator) {
		return E_POINTER;
	}

	Reset();
	return QueryInterface(IID_IUnknown, reinterpret_cast<LPVOID*>(ppEnumerator));
}


STDMETHODIMP StatusBarPanels::Add(BSTR panelText, LONG preferredWidth, PanelContentConstants content/* = pcText*/, AlignmentConstants alignment/* = alLeft*/, PanelBorderStyleConstants borderStyle/* = pbsSunken*/, LONG panelData/* = 0*/, LONG insertAt/* = -1*/, IStatusBarPanel** ppAddedPanel/* = NULL*/)
{
	ATLASSERT_POINTER(ppAddedPanel, IStatusBarPanel*);
	if(!ppAddedPanel) {
		return E_POINTER;
	}

	ATLASSERT(preferredWidth >= 0);
	if(preferredWidth < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	#ifdef USE_STL
		if(insertAt == -1) {
			insertAt = properties.panels.size();
		}
		ATLASSERT((insertAt >= 0) && (insertAt <= properties.panels.size()));
		if((insertAt < 0) || (insertAt > static_cast<LONG>(properties.panels.size()))) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	#else
		if(insertAt == -1) {
			insertAt = properties.panels.GetCount();
		}
		ATLASSERT((insertAt >= 0) && (insertAt <= static_cast<LONG>(properties.panels.GetCount())));
		if((insertAt < 0) || (insertAt > static_cast<LONG>(properties.panels.GetCount()))) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	#endif

	CComObject<StatusBarPanel>* pStatBarPanelObj = NULL;
	CComObject<StatusBarPanel>::CreateInstance(&pStatBarPanelObj);
	pStatBarPanelObj->AddRef();
	pStatBarPanelObj->SetOwner(properties.pOwnerStatBar);
	pStatBarPanelObj->properties.text = panelText;
	pStatBarPanelObj->properties.preferredWidth = preferredWidth;
	pStatBarPanelObj->properties.content = content;
	pStatBarPanelObj->properties.alignment = alignment;
	pStatBarPanelObj->properties.borderStyle = borderStyle;
	pStatBarPanelObj->properties.panelData = panelData;
	pStatBarPanelObj->Attach(insertAt);
	pStatBarPanelObj->QueryInterface(IID_IStatusBarPanel, reinterpret_cast<LPVOID*>(ppAddedPanel));

	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin();
		std::advance(it, insertAt);
		properties.panels.insert(it, pStatBarPanelObj);

		while(it != properties.panels.end()) {
			++((*it)->properties.panelIndex);
			++it;
		}
	#else
		POSITION p = properties.panels.FindIndex(insertAt);
		if(p) {
			p = properties.panels.InsertBefore(p, pStatBarPanelObj);
		} else {
			p = properties.panels.AddTail(pStatBarPanelObj);
		}
		POSITION buffer = p;

		properties.panels.GetNext(p);
		while(p) {
			++(properties.panels.GetAt(p)->properties.panelIndex);
			properties.panels.GetNext(p);
		}
	#endif
	properties.pOwnerStatBar->ApplyPanelWidths(FALSE);

	HWND hWndStatBar = properties.GetStatBarHWnd();
	ATLASSERT(IsWindow(hWndStatBar));

	BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
	if(simpleMode) {
		properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
		SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
	}

	// set the new panel's properties and shift the following panels
	#ifdef USE_STL
		it = properties.panels.begin();
		std::advance(it, insertAt);
		while(it != properties.panels.end()) {
			(*it)->ApplySettings(FALSE, FALSE);
			++it;
		}
	#else
		while(buffer) {
			properties.panels.GetAt(buffer)->ApplySettings(FALSE, FALSE);
			properties.panels.GetNext(buffer);
		}
	#endif

	if(simpleMode) {
		SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
		properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
	} else {
		InvalidateRect(hWndStatBar, NULL, TRUE);
	}

	m_bRequiresSave = TRUE;
	properties.pOwnerStatBar->SetDirty(TRUE);
	return S_OK;
}

STDMETHODIMP StatusBarPanels::Count(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef USE_STL
		*pValue = properties.panels.size();
	#else
		*pValue = properties.panels.GetCount();
	#endif
	return S_OK;
}

STDMETHODIMP StatusBarPanels::Remove(LONG panelIndex)
{
	#ifdef USE_STL
		if((panelIndex < 0) || (panelIndex >= static_cast<LONG>(properties.panels.size()))) {
			return DISP_E_BADINDEX;
		}
	#else
		if((panelIndex < 0) || (panelIndex >= static_cast<LONG>(properties.panels.GetCount()))) {
			return DISP_E_BADINDEX;
		}
	#endif

	// Are we removing the auto-sized panel?
	if(properties.pOwnerStatBar->properties.panelToAutoSize) {
		LONG l = 0;
		properties.pOwnerStatBar->properties.panelToAutoSize->get_Index(&l);
		if(l == panelIndex) {
			// yes
			properties.pOwnerStatBar->properties.panelToAutoSize = NULL;
		}
	}

	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::iterator it = properties.panels.begin();
		std::advance(it, panelIndex);
		std::list< CComObject<StatusBarPanel>* >::iterator buffer = it;

		while(++it != properties.panels.end()) {
			--((*it)->properties.panelIndex);
		}
		(*buffer)->Release();
		properties.panels.erase(buffer);
	#else
		POSITION p = properties.panels.FindIndex(panelIndex);
		POSITION buffer = p;
		properties.panels.GetNext(p);
		while(p) {
			--(properties.panels.GetAt(p)->properties.panelIndex);
			properties.panels.GetNext(p);
		}
		properties.panels.GetAt(buffer)->Release();
		properties.panels.RemoveAt(buffer);
	#endif
	properties.pOwnerStatBar->ApplyPanelWidths(FALSE);

	HWND hWndStatBar = properties.GetStatBarHWnd();
	ATLASSERT(IsWindow(hWndStatBar));

	BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
	if(simpleMode) {
		properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
		SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
	}

	#ifdef USE_STL
		if(properties.panels.size() == 0) {
	#else
		if(properties.panels.GetCount() == 0) {
	#endif
		// the last panel can't be removed, so set everything to NULL
		SendMessage(hWndStatBar, SB_SETTEXT, 0, NULL);
		SendMessage(hWndStatBar, SB_SETTIPTEXT, 0, NULL);
		SendMessage(hWndStatBar, SB_SETICON, 0, NULL);
	} else {
		// shift the following panels
		#ifdef USE_STL
			it = properties.panels.begin();
			std::advance(it, panelIndex);
			while(it != properties.panels.end()) {
				(*it)->ApplySettings(FALSE, FALSE);
				++it;
			}
		#else
			p = properties.panels.FindIndex(panelIndex);
			while(p) {
				properties.panels.GetAt(p)->ApplySettings(FALSE, FALSE);
				properties.panels.GetNext(p);
			}
		#endif
	}

	if(simpleMode) {
		SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
		properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
	} else {
		InvalidateRect(hWndStatBar, NULL, TRUE);
	}

	m_bRequiresSave = TRUE;
	properties.pOwnerStatBar->SetDirty(TRUE);
	return S_OK;
}

STDMETHODIMP StatusBarPanels::RemoveAll(void)
{
	#ifdef USE_STL
		std::list< CComObject<StatusBarPanel>* >::const_iterator it = properties.panels.begin();
		while(it != properties.panels.end()) {
			(*it)->Release();
			++it;
		}
		properties.panels.clear();
	#else
		POSITION p = properties.panels.GetHeadPosition();
		while(p) {
			properties.panels.GetAt(p)->Release();
			properties.panels.GetNext(p);
		}
		properties.panels.RemoveAll();
	#endif

	// we removed the auto-sized panel
	properties.pOwnerStatBar->properties.panelToAutoSize = NULL;

	HWND hWndStatBar = properties.GetStatBarHWnd();
	ATLASSERT(IsWindow(hWndStatBar));

	BOOL simpleMode = SendMessage(hWndStatBar, SB_ISSIMPLE, 0, 0);
	if(simpleMode) {
		properties.pOwnerStatBar->EnterInternalSBSIMPLESection();
		SendMessage(hWndStatBar, SB_SIMPLE, FALSE, 0);
	}

	// the last panel can't be removed, so set everything to NULL
	SendMessage(hWndStatBar, SB_SETTEXT, 0, NULL);
	SendMessage(hWndStatBar, SB_SETTIPTEXT, 0, NULL);
	SendMessage(hWndStatBar, SB_SETICON, 0, NULL);

	if(simpleMode) {
		SendMessage(hWndStatBar, SB_SIMPLE, TRUE, 0);
		properties.pOwnerStatBar->LeaveInternalSBSIMPLESection();
	}

	properties.pOwnerStatBar->ApplyPanelWidths(FALSE);

	m_bRequiresSave = TRUE;
	properties.pOwnerStatBar->SetDirty(TRUE);
	return S_OK;
}