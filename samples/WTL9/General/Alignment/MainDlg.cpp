// MainDlg.cpp : implementation of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#include <olectl.h>
#include "stdafx.h"
#include "resource.h"

#include "MainDlg.h"
#include ".\maindlg.h"


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

	for(int i = 0; i <= 2; i++) {
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

	controls.enablePanelCheckBox1 = GetDlgItem(IDC_ENABLE1CHECKBOX);
	controls.enablePanelCheckBox1.SetCheck(BST_CHECKED);
	controls.enablePanelCheckBox2 = GetDlgItem(IDC_ENABLE2CHECKBOX);
	controls.enablePanelCheckBox2.SetCheck(BST_CHECKED);
	controls.enablePanelCheckBox3 = GetDlgItem(IDC_ENABLE3CHECKBOX);
	controls.enablePanelCheckBox3.SetCheck(BST_CHECKED);
	controls.noteStatic = GetDlgItem(IDC_NOTESTATIC);
	controls.noteStatic.SetWindowText(TEXT("Note: To correctly align the icon, you should use tabulators instead of the Alignment property.\r\nSample:\r\nPanel.Text = vbTab && vbTab && \"Some Text\""));

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
	for(int i = 0; i <= 1; i++) {
		CComPtr<IStatusBarPanel> pPanel = pPanels->GetItem(i);
		pPanel->PuthIcon(HandleToLong(hPanelIcon[i + 1]));
	}
}

LRESULT CMainDlg::OnBnClickedShowproppages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComQIPtr<ISpecifyPropertyPages> pSPP = controls.statBarU;
		if(pSPP) {
			CAUUID pPages = {0};
			pSPP->GetPages(&pPages);
			CComPtr<IUnknown> pUnk = NULL;
			controls.statBarU->QueryInterface(IID_PPV_ARGS(&pUnk));
			OleCreatePropertyFrame(*this, 0, 0, OLESTR("StatBarU"), 1, &pUnk, pPages.cElems, pPages.pElems, 0, 0, NULL);
			CoTaskMemFree(pPages.pElems);
		}
	}
	return 0;
}

LRESULT CMainDlg::OnBnClickedEnable1checkbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
		CComPtr<IStatusBarPanel> pPanel = pPanels->GetItem(0);
		pPanel->PutEnabled((controls.enablePanelCheckBox1.GetCheck() == BST_CHECKED) ? VARIANT_TRUE : VARIANT_FALSE);
	}
	return 0;
}

LRESULT CMainDlg::OnBnClickedEnable2checkbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
		CComPtr<IStatusBarPanel> pPanel = pPanels->GetItem(1);
		pPanel->PutEnabled((controls.enablePanelCheckBox2.GetCheck() == BST_CHECKED) ? VARIANT_TRUE : VARIANT_FALSE);
	}
	return 0;
}

LRESULT CMainDlg::OnBnClickedEnable3checkbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
		CComPtr<IStatusBarPanel> pPanel = pPanels->GetItem(2);
		pPanel->PutEnabled((controls.enablePanelCheckBox3.GetCheck() == BST_CHECKED) ? VARIANT_TRUE : VARIANT_FALSE);
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

void __stdcall CMainDlg::ContextMenuStatbaru(LPDISPATCH panel, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/)
{
	CComQIPtr<IStatusBarPanel> pPanel = panel;
	if(pPanel) {
		if(pPanel->GetIndex() == 1) {
			HMENU hMenu = LoadMenu(ModuleHelper::GetResourceInstance(), MAKEINTRESOURCE(IDR_PANELMENU));
			HMENU hPopupMenu = GetSubMenu(hMenu, 0);

			switch(pPanel->GetAlignment()) {
				case alLeft:
					CheckMenuRadioItem(hPopupMenu, 0, 2, 0, MF_BYPOSITION);
					break;
				case alCenter:
					CheckMenuRadioItem(hPopupMenu, 0, 2, 1, MF_BYPOSITION);
					break;
				case alRight:
					CheckMenuRadioItem(hPopupMenu, 0, 2, 2, MF_BYPOSITION);
					break;
			}

			POINT pt = {x, y};
			::ClientToScreen(static_cast<HWND>(LongToHandle(controls.statBarU->GethWnd())), &pt);
			int selectedCmd = TrackPopupMenuEx(hPopupMenu, TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, static_cast<HWND>(LongToHandle(controls.statBarU->GethWnd())), NULL);

			switch(selectedCmd) {
				case ID_ALIGNMENT_LEFT:
					pPanel->PutAlignment(alLeft);
					controls.noteStatic.ShowWindow(SW_HIDE);
					break;
				case ID_ALIGNMENT_CENTER:
					pPanel->PutAlignment(alCenter);
					controls.noteStatic.ShowWindow(SW_SHOW);
					break;
				case ID_ALIGNMENT_RIGHT:
					pPanel->PutAlignment(alRight);
					controls.noteStatic.ShowWindow(SW_SHOW);
					break;
			}
			DestroyMenu(hPopupMenu);
			DestroyMenu(hMenu);
		}
	}
}

void __stdcall CMainDlg::RecreatedControlWindowStatbaru(long /*hWnd*/)
{
	SetupPanelIcons();
}
