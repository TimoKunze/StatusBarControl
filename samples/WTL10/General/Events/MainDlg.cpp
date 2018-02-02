// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"


BOOL CMainDlg::PreTranslateMessage(MSG* pMsg)
{
	return CWindow::IsDialogMessage(pMsg);
}

LRESULT CMainDlg::OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	CloseDialog(0);
	return 0;
}

LRESULT CMainDlg::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	if(controls.statBarU) {
		DispEventUnadvise(controls.statBarU);
		controls.statBarU.Release();
	}

	for(int i = 0; i <= 3; i++) {
		DestroyIcon(hPanelIcon[i]);
	}

	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);

	bHandled = FALSE;
	return 0;
}

LRESULT CMainDlg::OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	// Init resizing
	DlgResize_Init(false, false);

	// set icons
	HICON hIcon = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR));
	SetIcon(hIcon, TRUE);
	HICON hIconSmall = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_MAINFRAME), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	SetIcon(hIconSmall, FALSE);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);

	hPanelIcon[0] = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_PANELICON1), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	hPanelIcon[1] = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_PANELICON2), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	hPanelIcon[2] = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_PANELICON3), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
	hPanelIcon[3] = static_cast<HICON>(LoadImage(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDI_PANELICON4), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));

	controls.logEdit = GetDlgItem(IDC_EDITLOG);
	controls.simpleModeCheckbox = GetDlgItem(IDC_SIMPLEMODECHECKBOX);
	statBarUWnd.SubclassWindow(GetDlgItem(IDC_STATBARU));
	statBarUWnd.QueryControl(__uuidof(IStatusBar), reinterpret_cast<LPVOID*>(&controls.statBarU));
	if(controls.statBarU) {
		DispEventAdvise(controls.statBarU);
		SetupPanelIcons();
	}

	return TRUE;
}

LRESULT CMainDlg::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		controls.statBarU->About();
	}
	return 0;
}

LRESULT CMainDlg::OnBnClickedSimplemodecheckbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		controls.statBarU->PutSimpleMode((controls.simpleModeCheckbox.GetCheck() == BST_CHECKED) ? VARIANT_TRUE : VARIANT_FALSE);
	}
	return 0;
}

LRESULT CMainDlg::OnAxContainerWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.statBarU) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			if(IsZoomed()) {
				controls.statBarU->PutForceSizeGripperDisplay(VARIANT_FALSE);
			} else {
				controls.statBarU->PutForceSizeGripperDisplay(VARIANT_TRUE);
			}
		}
	}

	bHandled = FALSE;
	return 0;
}

void CMainDlg::AddLogEntry(CAtlString text)
{
	static int cLines = 0;
	static CAtlString oldText;

	cLines++;
	if(cLines > 50) {
		// delete the first line
		int pos = oldText.Find(TEXT("\r\n"));
		oldText = oldText.Mid(pos + lstrlen(TEXT("\r\n")), oldText.GetLength());
		cLines--;
	}
	oldText += text;
	oldText += TEXT("\r\n");

	controls.logEdit.SetWindowText(oldText);
	int l = oldText.GetLength();
	controls.logEdit.SetSel(l, l);
}

void CMainDlg::CloseDialog(int nVal)
{
	DestroyWindow();
	PostQuitMessage(nVal);
}

void CMainDlg::SetupPanelIcons(void)
{
	CComPtr<IStatusBarPanel> pSimplePanel = controls.statBarU->GetSimplePanel();
	pSimplePanel->PuthIcon(HandleToLong(hPanelIcon[0]));

	CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
	for(int i = 0; i <= 2; i++) {
		CComPtr<IStatusBarPanel> pPanel = pPanels->GetItem(i);
		pPanel->PuthIcon(HandleToLong(hPanelIcon[i + 1]));
	}
}

void __stdcall CMainDlg::ClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_Click: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_Click: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ContextMenuStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_ContextMenu: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_ContextMenu: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_DblClick: panel=");
		str += text;
		SysFreeString(text);

		PanelContentConstants content = pPanel->GetContent();
		switch(content) {
			case pcCapsLock:
			case pcInsertKey:
			case pcKanaLock:
			case pcNumLock:
			case pcScrollLock:
				pPanel->PutEnabled((pPanel->GetEnabled() != VARIANT_FALSE) ? VARIANT_FALSE : VARIANT_TRUE);
				break;
		}
	} else {
		str += TEXT("StatBarU_DblClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::DestroyedControlWindowStatbaru(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("StatBarU_DestroyedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
}

void __stdcall CMainDlg::MClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MClick: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MDblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MDblClick: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MDblClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseDownStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MouseDown: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MouseDown: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseEnterStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MouseEnter: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MouseEnter: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseHoverStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MouseHover: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MouseHover: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseLeaveStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MouseLeave: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MouseLeave: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseMoveStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MouseMove: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MouseMove: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::MouseUpStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_MouseUp: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_MouseUp: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragDropStatbaru(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("StatBarU_OLEDragDrop: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("StatBarU_OLEDragDrop: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<IStatusBarPanel> pPanel = dropTarget;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);

	if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
		_variant_t files = pData->GetData(CF_HDROP, -1, 1);
		if(files.vt == (VT_BSTR | VT_ARRAY)) {
			CComSafeArray<BSTR> array(files.parray);
			str = TEXT("");
			for(int i = array.GetLowerBound(); i <= array.GetUpperBound(); i++) {
				str += array.GetAt(i);
				str += TEXT("\r\n");
			}
			controls.statBarU->FinishOLEDragDrop();
			MessageBox(str, TEXT("Dropped files"));
		}
	}
}

void __stdcall CMainDlg::OLEDragEnterStatbaru(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("StatBarU_OLEDragEnter: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("StatBarU_OLEDragEnter: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<IStatusBarPanel> pPanel = dropTarget;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragLeaveStatbaru(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("StatBarU_OLEDragLeave: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("StatBarU_OLEDragLeave: data=NULL");
	}

	str += TEXT(", dropTarget=");
	CComQIPtr<IStatusBarPanel> pPanel = dropTarget;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::OLEDragMouseMoveStatbaru(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IOLEDataObject> pData = data;
	if(pData) {
		str += TEXT("StatBarU_OLEDragMouseMove: data=");
		if(pData->GetFormat(CF_HDROP, -1, 1) != VARIANT_FALSE) {
			_variant_t files = pData->GetData(CF_HDROP, -1, 1);
			if(files.vt == (VT_BSTR | VT_ARRAY)) {
				CComSafeArray<BSTR> array(files.parray);
				CAtlString tmp;
				tmp.Format(TEXT("%u files"), array.GetCount());
				str += tmp;
			} else {
				str += TEXT("<ERROR>");
			}
		} else {
			str += TEXT("0 files");
		}
	} else {
		str += TEXT("StatBarU_OLEDragMouseMove: data=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", effect=%i, dropTarget="), *effect);
	str += tmp;

	CComQIPtr<IStatusBarPanel> pPanel = dropTarget;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("NULL");
	}

	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::PanelMouseEnterStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_PanelMouseEnter: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_PanelMouseEnter: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::PanelMouseLeaveStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_PanelMouseLeave: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_PanelMouseLeave: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_RClick: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_RClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RDblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_RDblClick: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_RDblClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::RecreatedControlWindowStatbaru(long hWnd)
{
	CAtlString str;
	str.Format(TEXT("StatBarU_RecreatedControlWindow: hWnd=0x%X"), hWnd);

	AddLogEntry(str);
	SetupPanelIcons();
}

void __stdcall CMainDlg::ResizedControlWindowStatbaru()
{
	AddLogEntry(CAtlString(TEXT("StatBarU_ResizedControlWindow")));
}

void __stdcall CMainDlg::SetupToolTipWindowStatbaru(LPDISPATCH panel, long hWndToolTip)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_SetupToolTipWindow: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_SetupToolTipWindow: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", hWndToolTip=0x%X"), hWndToolTip);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::ToggledSimpleModeStatbaru()
{
	AddLogEntry(CAtlString(TEXT("StatBarU_ToggledSimpleMode")));

	controls.simpleModeCheckbox.SetCheck((controls.statBarU->GetSimpleMode() != VARIANT_FALSE) ? BST_CHECKED : BST_UNCHECKED);
}

void __stdcall CMainDlg::XClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_XClick: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_XClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}

void __stdcall CMainDlg::XDblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails)
{
	CAtlString str;
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		BSTR text = pPanel->GetText().Detach();
		str += TEXT("StatBarU_XDblClick: panel=");
		str += text;
		SysFreeString(text);
	} else {
		str += TEXT("StatBarU_XDblClick: panel=NULL");
	}
	CAtlString tmp;
	tmp.Format(TEXT(", button=%i, shift=%i, x=%i, y=%i, hitTestDetails=%i"), button, shift, x, y, hitTestDetails);
	str += tmp;

	AddLogEntry(str);
}
