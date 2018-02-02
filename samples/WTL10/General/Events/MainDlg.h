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
		CEdit logEdit;
		CButton simpleModeCheckbox;
		CComPtr<IStatusBar> statBarU;
	} controls;

	HICON hPanelIcon[4];

	virtual BOOL PreTranslateMessage(MSG* pMsg);

	BEGIN_MSG_MAP(CMainDlg)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_INITDIALOG, OnInitDialog)

		COMMAND_ID_HANDLER(ID_APP_ABOUT, OnAppAbout)

		COMMAND_HANDLER(IDC_SIMPLEMODECHECKBOX, BN_CLICKED, OnBnClickedSimplemodecheckbox)

		CHAIN_MSG_MAP(CDialogResize<CMainDlg>)

		ALT_MSG_MAP(1)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnAxContainerWindowPosChanged)
	END_MSG_MAP()

	BEGIN_SINK_MAP(CMainDlg)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 0, ClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 1, ContextMenuStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 2, DblClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 3, DestroyedControlWindowStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 4, MClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 5, MDblClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 6, MouseDownStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 7, MouseEnterStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 8, MouseHoverStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 9, MouseLeaveStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 10, MouseMoveStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 11, MouseUpStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 12, OLEDragDropStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 13, OLEDragEnterStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 14, OLEDragLeaveStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 15, OLEDragMouseMoveStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 17, PanelMouseEnterStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 18, PanelMouseLeaveStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 19, RClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 20, RDblClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 21, RecreatedControlWindowStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 22, ResizedControlWindowStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 23, SetupToolTipWindowStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 24, ToggledSimpleModeStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 25, XClickStatbaru)
		SINK_ENTRY_EX(IDC_STATBARU, __uuidof(_IStatusBarEvents), 26, XDblClickStatbaru)
	END_SINK_MAP()

	BEGIN_DLGRESIZE_MAP(CMainDlg)
		DLGRESIZE_CONTROL(IDC_STATBARU, DLSZ_SIZE_X | DLSZ_MOVE_Y)
		DLGRESIZE_CONTROL(IDC_EDITLOG, DLSZ_SIZE_X | DLSZ_SIZE_Y)
		DLGRESIZE_CONTROL(IDC_SIMPLEMODECHECKBOX, DLSZ_MOVE_X)
		DLGRESIZE_CONTROL(ID_APP_ABOUT, DLSZ_MOVE_X)
	END_DLGRESIZE_MAP()

	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled);
	LRESULT OnInitDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnBnClickedSimplemodecheckbox(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnAxContainerWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& bHandled);

	void CloseDialog(int nVal);
	void AddLogEntry(CAtlString text);
	void SetupPanelIcons(void);

	void __stdcall ClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall ContextMenuStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall DestroyedControlWindowStatbaru(long hWnd);
	void __stdcall MClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MDblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseDownStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseEnterStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseHoverStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseLeaveStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseMoveStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall MouseUpStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragDropStatbaru(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragEnterStatbaru(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragLeaveStatbaru(LPDISPATCH data, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall OLEDragMouseMoveStatbaru(LPDISPATCH data, long* effect, LPDISPATCH dropTarget, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall PanelMouseEnterStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall PanelMouseLeaveStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RDblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall RecreatedControlWindowStatbaru(long hWnd);
	void __stdcall ResizedControlWindowStatbaru();
	void __stdcall SetupToolTipWindowStatbaru(LPDISPATCH panel, long hWndToolTip);
	void __stdcall ToggledSimpleModeStatbaru();
	void __stdcall XClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
	void __stdcall XDblClickStatbaru(LPDISPATCH panel, short button, short shift, long x, long y, long hitTestDetails);
};
