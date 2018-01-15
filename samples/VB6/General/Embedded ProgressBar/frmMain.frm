VERSION 5.00
Object = "{A10D6B26-9A8F-4A87-A2D1-1D8C9EED0967}#1.5#0"; "StatBarU.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "StatusBar 1.5 - Embedded ProgressBar Sample"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdPropPages 
      Caption         =   "Show property pages..."
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto-sized panel"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'Kein
         Height          =   1600
         Left            =   100
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   180
         TabIndex        =   3
         Top             =   240
         Width           =   2700
         Begin VB.OptionButton optPanel3 
            Caption         =   "Panel 3"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optPanel2 
            Caption         =   "Panel 2"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton optPanel1 
            Caption         =   "Panel 1"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3240
      Top             =   1320
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
      Max             =   50
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin StatBarLibUCtl.StatusBar StatBarU 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      Top             =   2205
      Width           =   5895
      Version         =   256
      _cx             =   10398
      _cy             =   609
      Appearance      =   0
      BackColor       =   -1
      BorderStyle     =   0
      CustomCapsLockText=   "CAPS"
      CustomInsertKeyText=   "INS"
      CustomKanaLockText=   "KANA"
      CustomNumLockText=   "NUM"
      CustomScrollLockText=   "SCRL"
      DisabledEvents  =   7
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForceSizeGripperDisplay=   0   'False
      HoverTime       =   -1
      MinimumHeight   =   0
      MousePointer    =   0
      BeginProperty Panels {CCA75315-B100-4B5F-80F6-8DFE616F8FDB} 
         Version         =   256
         NumPanels       =   3
         BeginProperty Panel1 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   0
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   100
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   150
            RightToLeftText =   0   'False
            Text            =   "Ready"
            Object.ToolTipText     =   "Panel 1"
         EndProperty
         BeginProperty Panel2 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   0
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   100
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   100
            RightToLeftText =   0   'False
            Text            =   ""
            Object.ToolTipText     =   "Panel 2"
         EndProperty
         BeginProperty Panel3 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   1
            Enabled         =   0   'False
            ForeColor       =   -1
            MinimumWidth    =   100
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   70
            RightToLeftText =   0   'False
            Text            =   "CAPS"
            Object.ToolTipText     =   "Panel 3"
         EndProperty
      EndProperty
      PanelToAutoSize =   0
      RegisterForOLEDragDrop=   0   'False
      RightToLeftLayout=   0   'False
      ShowToolTips    =   -1  'True
      SimpleMode      =   0   'False
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      BeginProperty SimplePanel {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
         Version         =   256
         BorderStyle     =   1
         PanelData       =   0
         ParseTabs       =   -1  'True
         RightToLeftText =   0   'False
         Text            =   "Ready"
         Object.ToolTipText     =   "Simple Mode"
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As UUID) As Long
  Private Declare Function CoTaskMemFree Lib "ole32.dll" (ByVal cb As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Private Declare Function OleCreatePropertyFrame Lib "oleaut32.dll" (ByVal hWndOwner As Long, ByVal x As Long, ByVal y As Long, ByVal lpszCaption As Long, ByVal cObjects As Long, ByVal ppUnk As Long, ByVal cPages As Long, ByVal pPageClsID As Long, ByVal lcid As Long, ByVal dwReserved As Long, ByVal pvReserved As Long) As Long
  Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Private Sub cmdAbout_Click()
  StatBarU.About
End Sub

Private Sub cmdPropPages_Click()
  ShowProperties StatBarU, StatBarU.Name
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  SetParent ProgBar.hWnd, StatBarU.hWnd
  StatBarU_ResizedControlWindow
End Sub

Private Sub optPanel1_Click()
  Set StatBarU.PanelToAutoSize = StatBarU.Panels(0)
  StatBarU_ResizedControlWindow
End Sub

Private Sub optPanel2_Click()
  Set StatBarU.PanelToAutoSize = StatBarU.Panels(1)
  StatBarU_ResizedControlWindow
End Sub

Private Sub optPanel3_Click()
  Set StatBarU.PanelToAutoSize = StatBarU.Panels(2)
  StatBarU_ResizedControlWindow
End Sub

Private Sub StatBarU_ResizedControlWindow()
  Dim x1 As Long
  Dim x2 As Long
  Dim y1 As Long
  Dim y2 As Long

  StatBarU.Panels(1).GetRectangle x1, y1, x2, y2
  MoveWindow ProgBar.hWnd, x1, y1 + 1, x2 - x1 - 4, y2 - y1 - 2, 1
End Sub

Private Sub Timer1_Timer()
  ProgBar.Value = (ProgBar.Value + 1) Mod ProgBar.Max
End Sub


Private Sub ShowProperties(ByVal Ctrl As Object, Optional ByVal CtrlName As String, Optional ByVal hWndParent As Long = 0)
  Const strIID_ISpecifyPropertyPages = "{B196B28B-BAB4-101A-B69C-00AA00341D07}"
  Dim IID_ISpecifyPropertyPages As UUID
  Dim ISPP As IVBSpecifyPropertyPages
  Dim IUnk As IVBUnknown
  Dim propertyPages As CAUUID

  If hWndParent = 0 Then hWndParent = Me.hWnd
  CLSIDFromString StrPtr(strIID_ISpecifyPropertyPages), IID_ISpecifyPropertyPages

  Set IUnk = Ctrl.object
  IUnk.QueryInterface IID_ISpecifyPropertyPages, ISPP
  If Not (ISPP Is Nothing) Then
    ' get the CLSIDs of the pages
    ISPP.GetPages propertyPages
    ' show the pages
    OleCreatePropertyFrame hWndParent, 0, 0, StrPtr(CtrlName), 1, VarPtr(IUnk), propertyPages.cElems, propertyPages.pElems, 0, 0, 0
    ' release the pages array
    CoTaskMemFree propertyPages.pElems
    Set ISPP = Nothing
  End If
  Set IUnk = Nothing
End Sub
