VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBw 
   Caption         =   "Baldur's Woods"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar stbBw 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   3045
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.Frame fmeOpt 
      Caption         =   " 选项 "
      Height          =   2175
      Left            =   2040
      TabIndex        =   12
      Top             =   780
      Width           =   2895
      Begin VB.ComboBox cboClk 
         Height          =   315
         ItemData        =   "frmBw.frx":1782
         Left            =   240
         List            =   "frmBw.frx":17A7
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtCurPoints 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPoints 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblClk 
         AutoSize        =   -1  'True
         Caption         =   "每秒点击次数："
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label lblCurPoints 
         AutoSize        =   -1  'True
         Caption         =   "当前点数："
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   900
      End
      Begin VB.Label lblPoints 
         AutoSize        =   -1  'True
         Caption         =   "最低点数："
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   285
         Width           =   900
      End
   End
   Begin VB.Frame fmeGt 
      Caption         =   " 游戏种类 "
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   780
      Width           =   1815
      Begin MSComctlLib.ListView lvGt 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3201
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imglstGames"
         SmallIcons      =   "imglstGames"
         ForeColor       =   16777215
         BackColor       =   8421504
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "(&C)取消"
      Height          =   320
      Left            =   5040
      TabIndex        =   2
      Top             =   2640
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "(&S)查找"
      Default         =   -1  'True
      Height          =   320
      Left            =   5040
      Picture         =   "frmBw.frx":17D3
      TabIndex        =   0
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "(&O)确定"
      Height          =   320
      Left            =   5040
      TabIndex        =   1
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Frame fmeAbout 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   -240
      TabIndex        =   7
      Top             =   -120
      Width           =   6735
      Begin VB.PictureBox picGt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   375
         ScaleHeight     =   510
         ScaleWidth      =   495
         TabIndex        =   17
         Top             =   210
         Width           =   495
      End
      Begin VB.Image imgReadme 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6090
         Picture         =   "frmBw.frx":1B15
         Stretch         =   -1  'True
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2016, i0apps@gmail.com"
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.1"
         Height          =   195
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baldur's Woods"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   1290
      End
   End
   Begin MSComctlLib.ImageList imglstGames 
      Left            =   5400
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":2217
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":2AF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":33CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":4B5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":5437
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const KEY_0 = 48

Private Const KEY_9 = 57
Private Const KEY_BKSPACE = 8
Private Const MAIN_IMG = 3

Private clsBgw As New TbgWoods
Private Nid As NotifyIconData

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim gtValue As GameType

  On Error GoTo fndErrHandle
  stbBw.Panels(1).Text = "查找..."
  Me.MousePointer = 11
  clsBgw.SearchChrAttrs txtCurPoints.Text
  Me.MousePointer = 0
  stbBw.Panels(1).Text = "查找结果:"
  stbBw.Panels(2).Text = _
    "力量: " & CStr(clsBgw.Strength) & ", 敏捷: " & CStr(clsBgw.Dexterity) & _
    ", 体格: " & CStr(clsBgw.Constitution) & ", 智力: " & CStr(clsBgw.Intelligence) & _
    ", 智慧: " & CStr(clsBgw.Wisdom) & ", 魅力: " & CStr(clsBgw.Charisma)
fndErrHandle:
End Sub

Private Sub CmdOk_Click()
  ''On Error GoTo okErrHandle
  
  If (Len(txtPoints.Text) > 0 And Len(cboClk.Text) > 0) Then
    clsBgw.Points = CByte(txtPoints.Text)
    clsBgw.CriPoints = 0
    clsBgw.Delay = 1000 / CInt(cboClk.Text)
    clsBgw.CriStop = 0
    If EnableHotKeyHook Then
      clsBgw.SetBgData
      Nid.cbSize = Len(Nid)
      Nid.hWnd = picGt.hWnd
      Nid.uId = 1&
      Nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      Nid.ucallbackMessage = WM_MOUSEMOVE
      Nid.hIcon = picGt.Picture.Handle
      Nid.szTip = "Baldur's Woods" & Chr$(0)
      Shell_NotifyIcon NIM_ADD, Nid
      Me.Hide
      Me.WindowState = 0
      App.TaskVisible = False
    End If
  End If
''okErrHandle:
End Sub


Private Sub Form_Load()
Dim count As Single
Dim gtIcons(0 To 5) As Integer

  On Error GoTo loadErrHandle
  
  gtIcons(0) = 1
  gtIcons(1) = 1
  gtIcons(2) = 2
  gtIcons(3) = 3
  gtIcons(4) = 4
  gtIcons(5) = 5
  For count = 1 To 6
    lvGt.ListItems.Add count, clsBgw.TypeToStr(count - 1), clsBgw.GetName(count - 1), gtIcons(count - 1), gtIcons(count - 1)
  Next
  Set picGt.Picture = imglstGames.ListImages(2).Picture
  txtPoints.Text = "0"
  txtCurPoints.Text = "0"
  cboClk.ListIndex = 0
  
  defMinMax.MaxHeight = Me.Height
  defMinMax.MaxWidth = Me.Width
  defMinMax.MinHeight = Me.Height
  defMinMax.MinWidth = Me.Width
  
  oldWndProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
  SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf SubClass_WndMessage
loadErrHandle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SetWindowLong Me.hWnd, GWL_WNDPROC, oldWndProc
End Sub

Private Sub imgReadme_Click()
Dim lpParam As String
Dim lpDir As String

On Error GoTo imgErrHandle
lpParam = ""
lpDir = ""
  If FileExists(App.Path & "\readme.txt") Then
    Call ShellExecute(Me.hWnd, "open", App.Path & "\readme.txt", lpParam, lpDir, SW_SHOWNORMAL)
  End If
imgErrHandle:
End Sub

Private Sub lvGt_Click()
  If IsNull(lvGt.SelectedItem) = False Then
    Set picGt.Picture = imglstGames.ListImages(lvGt.SelectedItem.Icon).Picture
    stbBw.Panels(1).Text = "当前游戏:"
    stbBw.Panels(2).Text = lvGt.SelectedItem.Text
    clsBgw.GameType = CByte(lvGt.SelectedItem.Index - 1)
  End If
End Sub

Private Sub lvGt_ItemClick(ByVal Item As MSComctlLib.ListItem)
 Set picGt.Picture = imglstGames.ListImages(Item.Icon).Picture
 stbBw.Panels(1).Text = "当前游戏:"
 stbBw.Panels(2).Text = Item.Text
 clsBgw.GameType = CByte(Item.Index - 1)
End Sub

Private Sub picGt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result As Long

  If (Hex(x) = "1E1E") Then
    If (App.TaskVisible = False) Then
      Nid.cbSize = Len(Nid)
      Nid.hWnd = picGt.hWnd
      Nid.uId = 1&
      Shell_NotifyIcon NIM_DELETE, Nid
      App.TaskVisible = True
    End If
    DisableHotKeyHook
    Me.WindowState = 0
    Me.Show
    Result = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
    Result = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
  Else
    If (Hex(x) = "1E3C") Then
    End If
  End If
End Sub

Private Sub txtPoints_Change()
On Error Resume Next
  If (Len(txtPoints.Text) > 0) Then
    If CInt(txtPoints.Text) > clsBgw.MaxPoints Then
      txtPoints.Text = "0"
    End If
  End If
End Sub

Private Sub txtPoints_KeyPress(KeyAscii As Integer)
  If (KeyAscii < KEY_0 Or KeyAscii > KEY_9) And KeyAscii <> KEY_BKSPACE Then
    KeyAscii = 0
  End If
End Sub

Private Function FileExists(PathName As String) As Boolean
  If (Dir(PathName, vbDirectory) <> "") Then
    FileExists = True
  Else
    FileExists = False
  End If
End Function
