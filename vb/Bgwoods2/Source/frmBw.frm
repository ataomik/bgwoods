VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Baldur's Woods"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
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
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmCheck 
      Left            =   5640
      Top             =   2160
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   320
      Left            =   4830
      TabIndex        =   8
      Top             =   2760
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   320
      Left            =   4830
      TabIndex        =   7
      Top             =   1530
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   320
      Left            =   4830
      TabIndex        =   6
      Top             =   1065
      Width           =   1200
   End
   Begin VB.Frame fmeOptions 
      Caption         =   " Options "
      Height          =   2115
      Left            =   1800
      TabIndex        =   13
      Top             =   990
      Width           =   2790
      Begin VB.ComboBox cboCkTime 
         Height          =   315
         ItemData        =   "frmBw.frx":08CA
         Left            =   1560
         List            =   "frmBw.frx":08EF
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   930
         Width           =   1125
      End
      Begin VB.TextBox txtCurPoints 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   585
         Width           =   1125
      End
      Begin VB.TextBox txtPoints 
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   2580
         WordWrap        =   -1  'True
      End
      Begin VB.Line lnRight 
         BorderColor     =   &H00808080&
         X1              =   2700
         X2              =   2700
         Y1              =   255
         Y2              =   1260
      End
      Begin VB.Line lnLeft 
         BorderColor     =   &H00808080&
         X1              =   1530
         X2              =   1530
         Y1              =   255
         Y2              =   1260
      End
      Begin VB.Label lblCurPoints 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Points:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   105
         TabIndex        =   17
         Top             =   660
         Width           =   930
      End
      Begin VB.Label lblClick 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click time per second:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   135
         TabIndex        =   15
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Points:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   135
         TabIndex        =   14
         Top             =   315
         Width           =   420
      End
      Begin VB.Shape shpMsg 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H80000004&
         Height          =   360
         Left            =   60
         Top             =   1560
         Width           =   2640
      End
      Begin VB.Shape shpCkTime 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H80000004&
         Height          =   360
         Left            =   75
         Top             =   900
         Width           =   2640
      End
      Begin VB.Shape shpPoints 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H80000004&
         Height          =   360
         Left            =   75
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Frame fmeGames 
      Caption         =   " Games "
      Height          =   2130
      Left            =   15
      TabIndex        =   1
      Top             =   975
      Width           =   1650
      Begin MSComctlLib.ListView lvGames 
         Height          =   1815
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   3201
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   16777215
         BackColor       =   8421504
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ImageList imglstMain 
      Left            =   4830
      Top             =   2025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":091B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":11F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":1AD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":23AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":2C8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBw.frx":3567
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3210
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   450
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
   Begin VB.Frame fmeTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   -120
      TabIndex        =   0
      Top             =   -240
      Width           =   6375
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   330
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   18
         Top             =   420
         Width           =   615
      End
      Begin VB.Image imgHelp 
         Height          =   480
         Left            =   5655
         Picture         =   "frmBw.frx":3E43
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lblAbout 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2016, i0apps@gmail.com"
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   705
         Width           =   1785
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version 1.1"
         Height          =   195
         Left            =   2535
         TabIndex        =   11
         Top             =   420
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1065
         TabIndex        =   10
         Top             =   405
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BgRunlib As CBwLib
Private blvDragSignal As Boolean

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  With BgRunlib
    .Points = 0
    .CriPoints = 0
    .CriStop = 0
    .Delay = 1000 \ CInt(cboCkTime.Text)
    If Len(txtPoints.Text) > 0 Then
      .Points = CByte(txtPoints.Text)
    End If
  End With
  If CBwMod.EnableHotkey(Me.hWnd, BgRunlib) Then
    If BgRunlib.AddNotifyIcon(picTitle.Picture.Handle, picTitle.hWnd, "Baldur's woods", WM_MOUSEMOVE) Then
      Me.Visible = False
    End If
  End If
End Sub

Private Sub cmdSearch_Click()
  Dim S As String
  If Len(txtCurPoints.Text) > 1 Then
    BgRunlib.SetMemoryKeys txtCurPoints.Text
  End If
  With BgRunlib
    .DoSearch
    If .CustomKey Then
      S = "* "
    Else
      S = ""
    End If
    lblMsg.Caption = S & "Search Result: Str: " & CStr(.Strength) & ", Dex: " & CStr(.Dexterity) & ", Con: " & CStr(.Constitution) & ", Wis: " & CStr(.Wisdom) & ", Int: " & CStr(.Intelligence) & ", Chr: " & CStr(.Charisma) & ", Cri: " & CStr(.Crital)
  End With
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim gmTitles() As String, gmCaptions() As String
  Dim lstItem As ListItem
  
  Set BgRunlib = New CBwLib
  If IsObject(BgRunlib) Then
    BgRunlib.SetTimer tmCheck
  Else
    Unload Me
  End If
  
  Set picTitle.Picture = imglstMain.ListImages(1).ExtractIcon
  Set lvGames.Icons = imglstMain
  gmTitles = BgRunlib.Titles
  gmCaptions = BgRunlib.Captions
  For i = 0 To UBound(gmCaptions)
    Set lstItem = lvGames.ListItems.Add(i + 1, gmTitles(i), gmCaptions(i), i + 1)
    lstItem.ToolTipText = gmCaptions(i)
    lstItem.CreateDragImage
  Next
  StatusBar1.Panels(1).Text = "Current game:"
  StatusBar1.Panels(2).Text = gmCaptions(0)
  txtPoints.Text = "0"
  txtCurPoints.Text = "0"
  cboCkTime.ListIndex = 0
  
  With defMinMax
    .MinHeight = Me.Height
    .MaxHeight = Me.Height
    .MinWidth = Me.Width
    .MaxWidth = Me.Width
  End With
  oldWndProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf SubClass_WndMessage)
  Set Tm = tmCheck
  Tm.Interval = 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set BgRunlib = Nothing
  SetWindowLong Me.hWnd, GWL_WNDPROC, oldWndProc
End Sub

Private Sub imgHelp_Click()
  If Dir(App.Path & "\readme.txt") <> "" Then
    ShellExecute Me.hWnd, "Open", App.Path & "\readme.txt", vbNullString, App.Path, SW_NORMAL
  End If
End Sub

Private Sub lvGames_GotFocus()
  Dim i As Integer
  
  For i = 1 To lvGames.ListItems.Count
    If lvGames.ListItems(i).Ghosted = True Then
      lvGames.ListItems(i).Ghosted = False
    End If
  Next
End Sub

Private Sub lvGames_ItemClick(ByVal Item As MSComctlLib.ListItem)
  BgRunlib.SetDefalueValue Item.Index - 1
  If Len(txtCurPoints.Text) > 1 Then
    BgRunlib.SetMemoryKeys txtCurPoints.Text
  End If
  Set picTitle.Picture = imglstMain.ListImages(Item.Icon).ExtractIcon
  StatusBar1.Panels(1).Text = "Current game:"
  StatusBar1.Panels(2).Text = Item.Text
End Sub

Private Function ValueCheck(Text As String, num As Integer) As Boolean
  
  ValueCheck = False
  
  If Len(Text) > 2 Then
    If Len(Text) > 3 Or CInt(Left$(Text, 3) >= num) Then
      ValueCheck = True
    End If
  End If
End Function

Private Function InputCheck(Text As String) As String
  Dim c As String, S As String
  Dim i As Integer, j As Integer
  
  S = ""
  i = Len(Text)
  For j = 1 To i
    c = Mid$(Text, j, 1)
    If (Asc(c) > Asc("9")) Or (Asc(c) < Asc("0")) Then
    Else
      S = S & c
    End If
  Next
  
  InputCheck = S
End Function

Private Sub KeyPressCheck(KeyAscii As Integer)
  If ((KeyAscii < Asc("0")) Or (KeyAscii > Asc("9"))) And (KeyAscii <> 8) Then
    KeyAscii = 0
  End If
End Sub

Private Sub lvGames_LostFocus()
  If Not (lvGames.SelectedItem Is Nothing) Then
    lvGames.SelectedItem.Ghosted = True
  End If
End Sub

Private Sub lvGames_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  blvDragSignal = True
End Sub

Private Sub lvGames_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If Button = vbLeftButton Then
    If Not (lvGames.HitTest(x, y) Is Nothing) Then
      While blvDragSignal
        DoEvents
      Wend
    End If
  End If
End Sub

Private Sub lvGames_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  blvDragSignal = False
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Me.Visible = False And Button = vbLeftButton Then
    If BgRunlib.RemoveNotifyIcon Then
      CBwMod.DisableHotkey Me.hWnd
      Me.Visible = True
    End If
  End If
End Sub


Private Sub txtPoints_Change()
  Dim i As Integer
  Dim S As String
    
  i = txtPoints.SelStart
  S = InputCheck(txtPoints.Text)
  If Len(S) < Len(txtPoints.Text) Then
    i = i - 1
  End If
  If ValueCheck(S, BgRunlib.MaxPoints) Then
    S = CStr(BgRunlib.MaxPoints)
  End If
  txtPoints.Text = S
  txtPoints.SelStart = i
End Sub

Private Sub txtPoints_KeyPress(KeyAscii As Integer)
  If ValueCheck(txtPoints.Text, BgRunlib.MaxPoints) Then
    txtPoints.Text = CStr(BgRunlib.MaxPoints)
    If KeyAscii <> 8 And txtPoints.SelLength = 0 Then
      KeyAscii = 0
    End If
  Else
    KeyPressCheck KeyAscii
  End If
End Sub
