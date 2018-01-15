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

	KillTimer(10);

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

	controls.autoSizeRadioButton1 = GetDlgItem(IDC_AUTOSIZE1RADIOBUTTON);
	controls.autoSizeRadioButton1.SetCheck(BST_CHECKED);
	controls.autoSizeRadioButton2 = GetDlgItem(IDC_AUTOSIZE2RADIOBUTTON);
	controls.autoSizeRadioButton3 = GetDlgItem(IDC_AUTOSIZE3RADIOBUTTON);

	statBarUWnd.SubclassWindow(GetDlgItem(IDC_STATBARU));
	statBarUWnd.QueryControl(__uuidof(IStatusBar), reinterpret_cast<LPVOID*>(&controls.statBarU));
	if(controls.statBarU) {
		DispEventAdvise(controls.statBarU);
	}

	controls.progressBar = GetDlgItem(IDC_PROGRESS1);
	controls.progressBar.SetParent(static_cast<HWND>(LongToHandle(controls.statBarU->GethWnd())));
	ResizedControlWindowStatbaru();
	controls.progressBar.SetRange(0, 50);

	SetTimer(10, 250);

	return TRUE;
}

LRESULT CMainDlg::OnTimer(UINT /*uMsg*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	if(wParam == 10) {
		int minVal, maxVal;
		controls.progressBar.GetRange(minVal, maxVal);
		controls.progressBar.SetPos((controls.progressBar.GetPos() + 1) % maxVal);
	}
	return 0;
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

LRESULT CMainDlg::OnBnClickedAutosize1radiobutton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
		controls.statBarU->PutRefPanelToAutoSize(pPanels->GetItem(0));
		ResizedControlWindowStatbaru();
	}
	return 0;
}

LRESULT CMainDlg::OnBnClickedAutosize2radiobutton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
		controls.statBarU->PutRefPanelToAutoSize(pPanels->GetItem(1));
		ResizedControlWindowStatbaru();
	}
	return 0;
}

LRESULT CMainDlg::OnBnClickedAutosize3radiobutton(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	if(controls.statBarU) {
		CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
		controls.statBarU->PutRefPanelToAutoSize(pPanels->GetItem(2));
		ResizedControlWindowStatbaru();
	}
	return 0;
}

LRESULT CMainDlg::OnAxContainerWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled)
{
	if(controls.statBarU) {
		LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

		if((pDetails->flags & SWP_NOSIZE) == 0) {
			if(IsZoomed()) {
				if(controls.statBarU->GetForceSizeGripperDisplay() != VARIANT_FALSE) {
					// the control will be recreated
					controls.progressBar.SetParent(*this);
					controls.statBarU->PutForceSizeGripperDisplay(VARIANT_FALSE);
					controls.progressBar.SetParent(static_cast<HWND>(LongToHandle(controls.statBarU->GethWnd())));
					ResizedControlWindowStatbaru();
				}
			} else {
				if(controls.statBarU->GetForceSizeGripperDisplay() == VARIANT_FALSE) {
					// the control will be recreated
					controls.progressBar.SetParent(*this);
					controls.statBarU->PutForceSizeGripperDisplay(VARIANT_TRUE);
					controls.progressBar.SetParent(static_cast<HWND>(LongToHandle(controls.statBarU->GethWnd())));
					ResizedControlWindowStatbaru();
				}
			}
		}
	}

	bHandled = FALSE;
	return 0;
}

void __stdcall CMainDlg::ResizedControlWindowStatbaru()
{
	CComPtr<IStatusBarPanels> pPanels = controls.statBarU->GetPanels();
	CComPtr<IStatusBarPanel> pPanel = pPanels->GetItem(1);
	OLE_XPOS_PIXELS left, right;
	OLE_YPOS_PIXELS top, bottom;
	pPanel->GetRectangle(&left, &top, &right, &bottom);
	controls.progressBar.MoveWindow(left, top + 1, right - left - 4, bottom - top - 2, TRUE);
}
