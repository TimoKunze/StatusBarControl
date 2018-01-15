VERSION 5.00
Object = "{A10D6B26-9A8F-4A87-A2D1-1D8C9EED0967}#1.5#0"; "StatBarU.ocx"
Object = "{9262F940-61BD-45DD-A6A2-FF0054F6CF9C}#1.5#0"; "StatBarA.ocx"
Begin VB.Form frmMain 
   Caption         =   "StatusBar 1.5 - Events Sample"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
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
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   8760
      TabIndex        =   0
      Top             =   600
      Width           =   975
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
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtLog 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   0
      Width           =   8655
   End
   Begin StatBarLibACtl.StatusBar StatBarA 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      Top             =   5085
      Width           =   10545
      Version         =   256
      _cx             =   18600
      _cy             =   609
      Appearance      =   0
      BackColor       =   -1
      BorderStyle     =   0
      CustomCapsLockText=   ""
      CustomInsertKeyText=   ""
      CustomKanaLockText=   ""
      CustomNumLockText=   "NUM"
      CustomScrollLockText=   ""
      DisabledEvents  =   0
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
      BeginProperty Panels {9A38B78A-CE98-440B-B1F6-1CE994432F2B} 
         Version         =   256
         NumPanels       =   3
         BeginProperty Panel1 {72208630-41D0-479F-B075-7727847698DE} 
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
            Text            =   "Ready"
            Object.ToolTipText     =   "Panel 1"
         EndProperty
         BeginProperty Panel2 {72208630-41D0-479F-B075-7727847698DE} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   2
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   60
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   60
            RightToLeftText =   0   'False
            Text            =   "NUM"
            Object.ToolTipText     =   "Panel 2"
         EndProperty
         BeginProperty Panel3 {72208630-41D0-479F-B075-7727847698DE} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   5
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   80
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   80
            RightToLeftText =   0   'False
            Text            =   "EINFG"
            Object.ToolTipText     =   "Panel 3"
         EndProperty
      EndProperty
      PanelToAutoSize =   0
      RegisterForOLEDragDrop=   -1  'True
      RightToLeftLayout=   0   'False
      ShowToolTips    =   -1  'True
      SimpleMode      =   0   'False
      SupportOLEDragImages=   -1  'True
      UseSystemFont   =   -1  'True
      BeginProperty SimplePanel {72208630-41D0-479F-B075-7727847698DE} 
         Version         =   256
         BorderStyle     =   1
         PanelData       =   0
         ParseTabs       =   -1  'True
         RightToLeftText =   0   'False
         Text            =   "Ready"
         Object.ToolTipText     =   "Simple Mode"
      EndProperty
   End
   Begin StatBarLibUCtl.StatusBar StatBarU 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      Top             =   5430
      Width           =   10545
      Version         =   256
      _cx             =   18600
      _cy             =   609
      Appearance      =   0
      BackColor       =   -1
      BorderStyle     =   0
      CustomCapsLockText=   ""
      CustomInsertKeyText=   ""
      CustomKanaLockText=   ""
      CustomNumLockText=   "NUM"
      CustomScrollLockText=   ""
      DisabledEvents  =   0
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
            Text            =   "Ready"
            Object.ToolTipText     =   "Panel 1"
         EndProperty
         BeginProperty Panel2 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   2
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   60
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   60
            RightToLeftText =   0   'False
            Text            =   "NUM"
            Object.ToolTipText     =   "Panel 2"
         EndProperty
         BeginProperty Panel3 {CB0F173F-9E1F-4365-BF3C-6CC52F8C268B} 
            Version         =   256
            Alignment       =   0
            BorderStyle     =   0
            Content         =   5
            Enabled         =   -1  'True
            ForeColor       =   -1
            MinimumWidth    =   80
            PanelData       =   0
            ParseTabs       =   -1  'True
            PreferredWidth  =   80
            RightToLeftText =   0   'False
            Text            =   "EINFG"
            Object.ToolTipText     =   "Panel 3"
         EndProperty
      EndProperty
      PanelToAutoSize =   0
      RegisterForOLEDragDrop=   -1  'True
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

  Private bLog As Boolean
  Private hIcon(0 To 5) As Long


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  StatBarU.About
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
  While iconPath <> "" And i < 6
    hIcon(i) = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    iconPath = Dir$
    i = i + 1
  Wend
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked

  SetupPanelIconsA
  SetupPanelIconsU
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    On Error Resume Next
    cmdAbout.Left = Me.ScaleWidth - 5 - cmdAbout.Width
    chkLog.Left = cmdAbout.Left
    txtLog.Width = cmdAbout.Left - 5 - txtLog.Left
    txtLog.Height = StatBarA.Top - 2 - txtLog.Top
  End If
End Sub

Private Sub Form_Terminate()
  Dim i As Long

  For i = 0 To 5
    DestroyIcon hIcon(i)
  Next i
End Sub

Private Sub StatBarA_Click(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_Click: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_Click: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_ContextMenu(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_ContextMenu: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_ContextMenu: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_DblClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  Select Case panel.Content
    Case PanelContentConstants.pcCapsLock, PanelContentConstants.pcInsertKey, PanelContentConstants.pcKanaLock, PanelContentConstants.pcNumLock, PanelContentConstants.pcScrollLock
      panel.Enabled = Not panel.Enabled
  End Select
  If panel Is Nothing Then
    AddLogEntry "StatBarA_DblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_DblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "StatBarA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub StatBarA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "StatBarA_DragDrop"
End Sub

Private Sub StatBarA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "StatBarA_DragOver"
End Sub

Private Sub StatBarA_MClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_MDblClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MDblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MDblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_MouseDown(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MouseDown: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MouseDown: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_MouseEnter(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MouseEnter: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MouseEnter: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_MouseHover(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MouseHover: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MouseHover: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_MouseLeave(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MouseLeave: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MouseLeave: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_MouseMove(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
'  If panel Is Nothing Then
'    AddLogEntry "StatBarA_MouseMove: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  Else
'    AddLogEntry "StatBarA_MouseMove: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  End If
End Sub

Private Sub StatBarA_MouseUp(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_MouseUp: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_MouseUp: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_OLEDragDrop(ByVal data As StatBarLibACtl.IOLEDataObject, effect As StatBarLibACtl.OLEDropEffectConstants, ByVal dropTarget As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    StatBarA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub StatBarA_OLEDragEnter(ByVal data As StatBarLibACtl.IOLEDataObject, effect As StatBarLibACtl.OLEDropEffectConstants, ByVal dropTarget As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub StatBarA_OLEDragLeave(ByVal data As StatBarLibACtl.IOLEDataObject, ByVal dropTarget As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub StatBarA_OLEDragMouseMove(ByVal data As StatBarLibACtl.IOLEDataObject, effect As StatBarLibACtl.OLEDropEffectConstants, ByVal dropTarget As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub StatBarA_OwnerDrawPanel(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal hDC As Long, drawingRectangle As StatBarLibACtl.RECTANGLE)
  Dim str As String

  str = "StatBarA_OwnerDrawPanel: panel="
  If panel Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & panel.Text
  End If
  str = str & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub StatBarA_PanelMouseEnter(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_PanelMouseEnter: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_PanelMouseEnter: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_PanelMouseLeave(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_PanelMouseLeave: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_PanelMouseLeave: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_RClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_RClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_RClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_RDblClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_RDblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_RDblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "StatBarA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  SetupPanelIconsA
End Sub

Private Sub StatBarA_ResizedControlWindow()
  AddLogEntry "StatBarA_ResizedControlWindow"
  txtLog.Height = StatBarA.Top - 2 - txtLog.Top
End Sub

Private Sub StatBarA_SetupToolTipWindow(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal hWndToolTip As Long)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_SetupToolTipWindow: panel=Nothing, hWndToolTip=0x" & Hex(hWndToolTip)
  Else
    AddLogEntry "StatBarA_SetupToolTipWindow: panel=" & panel.Text & ", hWndToolTip=0x" & Hex(hWndToolTip)
  End If
End Sub

Private Sub StatBarA_ToggledSimpleMode()
  AddLogEntry "StatBarA_ToggledSimpleMode"
End Sub

Private Sub StatBarA_XClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_XClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_XClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarA_XDblClick(ByVal panel As StatBarLibACtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibACtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarA_XDblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarA_XDblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_Click(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_Click: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_Click: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_ContextMenu(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_ContextMenu: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_ContextMenu: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_DblClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  Select Case panel.Content
    Case PanelContentConstants.pcCapsLock, PanelContentConstants.pcInsertKey, PanelContentConstants.pcKanaLock, PanelContentConstants.pcNumLock, PanelContentConstants.pcScrollLock
      panel.Enabled = Not panel.Enabled
  End Select
  If panel Is Nothing Then
    AddLogEntry "StatBarU_DblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_DblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "StatBarU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub StatBarU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "StatBarU_DragDrop"
End Sub

Private Sub StatBarU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "StatBarU_DragOver"
End Sub

Private Sub StatBarU_MClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_MDblClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MDblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MDblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_MouseDown(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MouseDown: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MouseDown: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_MouseEnter(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MouseEnter: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MouseEnter: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_MouseHover(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MouseHover: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MouseHover: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_MouseLeave(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MouseLeave: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MouseLeave: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_MouseMove(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
'  If panel Is Nothing Then
'    AddLogEntry "StatBarU_MouseMove: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  Else
'    AddLogEntry "StatBarU_MouseMove: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  End If
End Sub

Private Sub StatBarU_MouseUp(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_MouseUp: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_MouseUp: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_OLEDragDrop(ByVal data As StatBarLibUCtl.IOLEDataObject, effect As StatBarLibUCtl.OLEDropEffectConstants, ByVal dropTarget As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    StatBarU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub StatBarU_OLEDragEnter(ByVal data As StatBarLibUCtl.IOLEDataObject, effect As StatBarLibUCtl.OLEDropEffectConstants, ByVal dropTarget As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub StatBarU_OLEDragLeave(ByVal data As StatBarLibUCtl.IOLEDataObject, ByVal dropTarget As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub StatBarU_OLEDragMouseMove(ByVal data As StatBarLibUCtl.IOLEDataObject, effect As StatBarLibUCtl.OLEDropEffectConstants, ByVal dropTarget As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "StatBarU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub StatBarU_OwnerDrawPanel(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal hDC As Long, drawingRectangle As StatBarLibUCtl.RECTANGLE)
  Dim str As String

  str = "StatBarU_OwnerDrawPanel: panel="
  If panel Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & panel.Text
  End If
  str = str & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"

  AddLogEntry str
End Sub

Private Sub StatBarU_PanelMouseEnter(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_PanelMouseEnter: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_PanelMouseEnter: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_PanelMouseLeave(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_PanelMouseLeave: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_PanelMouseLeave: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_RClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_RClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_RClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_RDblClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_RDblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_RDblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "StatBarU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  SetupPanelIconsU
End Sub

Private Sub StatBarU_ResizedControlWindow()
  AddLogEntry "StatBarU_ResizedControlWindow"
End Sub

Private Sub StatBarU_SetupToolTipWindow(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal hWndToolTip As Long)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_SetupToolTipWindow: panel=Nothing, hWndToolTip=0x" & Hex(hWndToolTip)
  Else
    AddLogEntry "StatBarU_SetupToolTipWindow: panel=" & panel.Text & ", hWndToolTip=0x" & Hex(hWndToolTip)
  End If
End Sub

Private Sub StatBarU_ToggledSimpleMode()
  AddLogEntry "StatBarU_ToggledSimpleMode"
End Sub

Private Sub StatBarU_XClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_XClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_XClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub StatBarU_XDblClick(ByVal panel As StatBarLibUCtl.IStatusBarPanel, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As StatBarLibUCtl.HitTestConstants)
  If panel Is Nothing Then
    AddLogEntry "StatBarU_XDblClick: panel=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "StatBarU_XDblClick: panel=" & panel.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

Private Sub SetupPanelIconsA()
  With StatBarA.Panels
    .Item(0).hIcon = hIcon(0)
    .Item(1).hIcon = hIcon(1)
    .Item(2).hIcon = hIcon(2)
  End With
End Sub

Private Sub SetupPanelIconsU()
  With StatBarU.Panels
    .Item(0).hIcon = hIcon(3)
    .Item(1).hIcon = hIcon(4)
    .Item(2).hIcon = hIcon(5)
  End With
End Sub
