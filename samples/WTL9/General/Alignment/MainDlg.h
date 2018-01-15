// MainDlg.h : interface of the CMainDlg class
//
/////////////////////////////////////////////////////////////////////////////

#pragma once
#import <libid:A10D6B26-9A8F-4a87-A2D1-1D8C9EED0967> version("1.5") named_guids, no_namespace, raw_dispinterfaces

class CMainDlg :
    public CAxDialogImpl<CMainDlg>,
    public CMessageFilter,
    public CDialogResize<CMainDlg>,
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CMainDlg>,
    public IDispEventImpl<IDC_STATBARU, CMainDlg, &__uuidof(_IStatusBarEvents), &LIBID_StatBarLibU, 1, 5>
{
public:
	enum { IDD = IDD_MAINDLG };

	CContainedWindowT<CAxWindow> statBarUWnd;

	CMainDlg() : statBarUWnd(this, 1)
	{
	}

	struct Controls
	{
		CButton enablePanelCheckBox1;
		CButton enablePanelCheckBox2;
		CButton enablePanelCheckBox3;
		CStatic noteStatic;
		CComPtr<IStatusBar> statBarU;
	} controls;

	HICON hPanelIcon[3];

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		COMMAND_HANDLER(IDC_SHOWPROPPAGES, BN_CLICKED, OnBnClickedShowproppages)
		COMMAND_HANDLER(IDC_ENABLE1CHECKBOX, BN_CLICKED, OnBnClickedEnable1checkbox)
		COMMAND_HANDLER(IDC_ENABLE2CHECKBOX, BN_CLICKED, OnBnClickedEnable2checkbox)
		COMMAND_HANDLER(IDC_ENABLE3CHECKBOX, BN_CLICKED, OnBnClickedEnable3checkbox)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnAxContainerWindowPosChanged)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 1, ContextMenuStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 21, RecreatedControlWindowStatbaru)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_STATBARU, DLSZ_SIZE_X | DLSZ_MOVE_Y)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedShowproppages(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedEnable1checkbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedEnable2checkbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedEnable3checkbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAxContainerWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);

	void CloseDialog(int nVal);
	void SetupPanelIcons(void);

	void __stdcall ContextMenuStatbaru(LPDISPATCH panel, short /*button*/, short /*shift*/, long x, long y, long /*hitTestDetails*/);
	void __stdcall RecreatedControlWindowStatbaru(long /*hWnd*/);
};
