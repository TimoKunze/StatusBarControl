VERSION 5.00
Object = "{A10D6B26-9A8F-4A87-A2D1-1D8C9EED0967}#1.5#0"; "StatBarU.ocx"
Begin VB.Form frmMain 
   Caption         =   "StatusBar 1.5 - Alignment Sample"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6120
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
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdPropPages 
      Caption         =   "Show property pages..."
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enable Panel 3"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Aktiviert
      Width           =   1380
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enable Panel 2"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Aktiviert
      Width           =   1380
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enable Panel 1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Multi-column panels shouldn't be disabled (results in drawing bugs)."
      Top             =   720
      Value           =   1  'Aktiviert
      Width           =   1380
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin StatBarLibUCtl.StatusBar StatBarU 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      Top             =   3150
      Width           =   6120
      Version         =   256
      _cx             =   10795
      _cy             =   609
      Appearance      =   0
      BackColor       =   -1
      BorderStyle     =   0
      CustomCapsLockText=   ""
      CustomInsertKeyText=   ""
      CustomKanaLockText=   ""
      CustomNumLockText=   ""
      CustomScrollLockText=   ""
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
            PreferredWidth  =   200
            RightToLeftText =   0   'False
            Text            =   "Col 1	Col 2	Col 3"
            Object.ToolTipText     =   "Using tabulators it's easy to have multiple columns in the same panel."
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
            PreferredWidth  =   150
            RightToLeftText =   0   'False
            Text            =   "Some Text"
            Object.ToolTipText     =   "Right-click this panel to select the alignment."
         EndProperty
         BeginProperty Panel3 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   256
            Alignment       =   2
            BorderStyle     =   0
            Content         =   7
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   100
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   80
            RightToLeftText =   0   'False
            Text            =   "19:43:42"
            Object.ToolTipText     =   "Current time"
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
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Menu mnuAlignment 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAlignmentLeft 
         Caption         =   "&Left"
      End
      Begin VB.Menu mnuAlignmentCenter 
         Caption         =   "&Center"
      End
      Begin VB.Menu mnuAlignmentRight 
         Caption         =   "&Right"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private hIcon(0 To 1) As Long


  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As UUID) As Long
  Private Declare Function CoTaskMemFree Lib "ole32.dll" (ByVal cb As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function OleCreatePropertyFrame Lib "oleaut32.dll" (ByVal hWndOwner As Long, ByVal x As Long, ByVal y As Long, ByVal lpszCaption As Long, ByVal cObjects As Long, ByVal ppUnk As Long, ByVal cPages As Long, ByVal pPageClsID As Long, ByVal lcid As Long, ByVal dwReserved As Long, ByVal pvReserved As Long) As Long


Private Sub chkEnabled_Click(Index As Integer)
  StatBarU.Panels(Index).Enabled = (chkEnabled(Index).Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  StatBarU.About
End Sub

Private Sub cmdPropPages_Click()
  ShowProperties StatBarU, StatBarU.Name
End Sub

Private Sub Form_Initialize()
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim i As Long
  Dim iconsDir As String
  Dim iconPath As String

  InitCommonControls

  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconPath = Dir$(iconsDir & "*.ico")
  i = 0
  While iconPath <> "" And i < 2
    hIcon(i) = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    iconPath = Dir$
    i = i + 1
  Wend
End Sub

Private Sub Form_Load()
  lblNote.Caption = "Note: To correctly align the icon, you should use tabulators instead of the Alignment property." & vbNewLine & "Sample:" & vbNewLine & "Panel.Text = vbTab & vbTab & ""Some Text"""

  SetupPanelIconsU
End Sub

Private Sub Form_Terminate()
  Dim i As Long

  For i = 0 To 1
    DestroyIcon hIcon(i)
  Next i
End Sub

Private Sub mnuAlignmentCenter_Click()
  StatBarU.Panels(1).Alignment = StatBarLibUCtl.AlignmentConstants.alCenter
  lblNote.Visible = True
End Sub

Private Sub mnuAlignmentLeft_Click()
  StatBarU.Panels(1).Alignment = StatBarLibUCtl.AlignmentConstants.alLeft
  lblNote.Visible = False
End Sub

Private Sub mnuAlignmentRight_Click()
  StatBarU.Panels(1).Alignment = StatBarLibUCtl.AlignmentConstants.alRight
  lblNote.Visible = True
End Sub

Private Sub StatBarU_ContextMenu(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel.Index = 1 Then
    mnuAlignmentLeft.Checked = False
    mnuAlignmentCenter.Checked = False
    mnuAlignmentRight.Checked = False
    Select Case panel.Alignment
      Case StatBarLibUCtl.AlignmentConstants.alLeft
        mnuAlignmentLeft.Checked = True
      Case StatBarLibUCtl.AlignmentConstants.alCenter
        mnuAlignmentCenter.Checked = True
      Case StatBarLibUCtl.AlignmentConstants.alRight
        mnuAlignmentRight.Checked = True
    End Select

    PopupMenu mnuAlignment, , StatBarU.Left + Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels), StatBarU.Top + Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
  End If
End Sub

Private Sub StatBarU_RecreatedControlWindow(ByVal hWnd As Long)
  SetupPanelIconsU
End Sub


Private Sub SetupPanelIconsU()
  With StatBarU.Panels
    .Item(0).hIcon = hIcon(0)
    .Item(1).hIcon = hIcon(1)
  End With
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
