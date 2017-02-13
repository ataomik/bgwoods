VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBwLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CGameType
  BG1 = 0
  TSC = 1
  SOA = 2
  TOB = 3
  IWD = 4
  HOW = 5
End Enum

Public Enum CAttrType
  Str = 0
  Cri = 1
  Inl = 2
  Wis = 3
  Dex = 4
  Con = 5
  Cha = 6
End Enum

Private Type CChrAttrs
  Str As Byte
  Dex As Byte
  Con As Byte
  Int As Byte
  Wis As Byte
  Cha As Byte
  Cri As Byte
End Type

Private Type CBwData
  GameType As Byte
  Points As Byte
  CriPoints As Byte
  Delay As Long
  CriStop As Byte
  Address As Long
  KeyStart As Integer
  KeyStop As Integer
  scrX As Integer
  scrY As Integer
End Type

Private Type CBwItem
  GameType As Byte
  ClassName As String
  WindowName As String
  MemoryKeys() As Byte
  BeginAddress As Long
  EndAddress As Long
  Offset As Integer
  CustomKey As Boolean
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type NotifyIconData
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private Type CNotifyIconData
  Nid As NotifyIconData
  Status As Long
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NotifyIconData) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesRead As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, lpFileMappigAttributes As SECURITY_ATTRIBUTES, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF
Private Const SECTION_QUERY = &H1
Private Const SECTION_MAP_WRITE = &H2
Private Const SECTION_MAP_READ = &H4
Private Const SECTION_MAP_EXECUTE = &H8
Private Const SECTION_EXTEND_SIZE = &H10
Private Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Private Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS
Private Const PAGE_READWRITE = &H4
Private Const MEF_LEFTDOWN = &H2
Private Const MEF_LEFTUP = &H4

Private Const BUFF_SIZE = 4096
Private Const CHRATTRS_COUNT = 7
Private Const GAME_NUM = 6

Private Const NameOfMapFile = "CBgWoods"

Private Const gmTitle = "Bg1,Tsc,Soa,Tob,Iwd,How"
Private Const gmCaption = "Baldur's Gate I,Baldur's Gate I - Tales of the Sword Coast,Baldur's Gate II - Shadows of Amn,Baldur's Gate II - Throne of Bhaal,Icewind Dale I,Icewind Dale I - Heart of Winter"
Private Const gmClsName = "ChitinClass"
Private Const gmEEClsName = "SDL_app"
Private Const gmKeyBg1 = "&H1E,&H36,&H34,&HC,&H17,&H1E"
Private Const gmKeySoa = "&H43,&H48,&H41,&H52,&H42,&H41,&H53,&H45"
Private Const gmKeyIwd = "&H1E,&H27,&H29,&HC,&H17,&H1E"
Private Const gmOffset = "524,524,590,590,524,524"
Private Const gmBeginAddr = 33554432
Private Const gmEndAddr = 201326592
Private Const keyNumStar = 106
Private Const KeyNumSlash = 111

Private Const MAX_POINTS = 108
Private Const MAX_CRIPOINTS = 100

Private p_Nid As CNotifyIconData
Private p_hProcess As Long
Private p_hMap As Long
Private p_BwItem As CBwItem
Private p_BwData As CBwData
Private p_ChrAttrs As CChrAttrs
Private p_bRun As Boolean
Private WithEvents p_Timer As Timer
Attribute p_Timer.VB_VarHelpID = -1

Public Function GetWindow(lpClassName As String, lpWindowName As String) As Long
  If lpWindowName <> "" Then
    GetWindow = FindWindow(lpClassName, lpWindowName)
  Else
    GetWindow = FindWindow(lpClassName, vbNullString)
  End If
End Function

Public Function GetThreadId(lpClassName As String, lpWindowName As String) As Long
  Dim hGame, threadId, pId As Long
    
  GetThreadId = 0
  If lpWindowName <> "" Then
    hGame = FindWindow(lpClassName, lpWindowName)
  Else
    hGame = FindWindow(lpClassName, vbNullString)
  End If
  If hGame <> 0 Then
    threadId = GetWindowThreadProcessId(hGame, pId)
    If threadId <> 0 Then
      GetThreadId = threadId
    End If
  End If
End Function

Public Function GetProcess(lpClassName As String, lpWindowName As String) As Long
  Dim hGame, threadId, pId, hProcess  As Long
  
  GetProcess = 0
  If lpWindowName <> "" Then
    hGame = FindWindow(lpClassName, lpWindowName)
  Else
    hGame = FindWindow(lpClassName, vbNullString)
  End If
  If hGame <> 0 Then
    threadId = GetWindowThreadProcessId(hGame, pId)
    If threadId <> 0 Then
      hProcess = OpenProcess(PROCESS_ALL_ACCESS, 1, pId)
      If hProcess <> 0 Then
        GetProcess = hProcess
      End If
    End If
  End If
End Function

Public Function Search(hProcess As Long, BeginAddr As Long, EndAddr As Long, lpData() As Byte) As Long
  Dim dtLen As Integer
  Dim Count As Long, lpNOBR As Long, Pos As Long
  Dim lpBuff() As Byte
  
  Search = 0
  On Error GoTo sechHandler
  If lpData(0) = 0 Then
    Exit Function
  End If
  dtLen = UBound(lpData)
  ReDim lpBuff(0 To BUFF_SIZE + dtLen - 2)
  
  Count = BeginAddr
  While Count < EndAddr
    If ReadProcessMemory(hProcess, Count, lpBuff(0), BUFF_SIZE + dtLen - 1, lpNOBR) Then
      Pos = InStrB(1, lpBuff, lpData, vbBinaryCompare)
      If Pos > 0 Then
        Search = Count + Pos - 1
        Count = EndAddr
      End If
    End If
    Count = Count + BUFF_SIZE
  Wend
  Exit Function
  
sechHandler:
End Function

Public Function OpenShareData(MapFileName As String, Ptr As Long, dwLength As Long, lpNameOfMapFile As String) As Long
  Dim st As SECURITY_ATTRIBUTES
  Dim hMap As Long
  
  hMap = 0
  st.bInheritHandle = 0
  st.lpSecurityDescriptor = 0
  st.nLength = 0
  hMap = CreateFileMapping(&HFFFFFFFF, st, PAGE_READWRITE, 0, dwLength, lpNameOfMapFile)
  If hMap <> 0 Then
    Ptr = MapViewOfFile(hMap, FILE_MAP_ALL_ACCESS, 0, 0, dwLength)
    If Ptr = 0 Then
      CloseHandle hMap
      hMap = 0
    End If
  End If
  
  OpenShareData = hMap
End Function

Public Sub CloseShareData(hMap As Long, Ptr As Long)
  If hMap <> 0 Then
    If Ptr <> 0 Then
      UnmapViewOfFile Ptr
    End If
    CloseHandle hMap
    hMap = 0
  End If
End Sub

Private Sub Class_Initialize()
  p_BwItem.GameType = BG1
  p_BwItem.ClassName = gmClsName
  p_BwItem.WindowName = ""
  p_BwItem.MemoryKeys = GetMemoryKeys(BG1)
  p_BwItem.BeginAddress = gmBeginAddr
  p_BwItem.EndAddress = gmEndAddr
  p_BwItem.Offset = GetOffSet(BG1)
  p_BwItem.CustomKey = False
  p_BwData.KeyStart = keyNumStar
  p_BwData.KeyStop = KeyNumSlash
  Set p_Timer = Nothing
End Sub

Private Function CheckAttrs() As Long
  Dim i As Integer
  
  i = 0
  Do While p_bRun
    If i > 5 Then
      DoEvents
      i = 0
    Else
        i = i + 1
    End If
    
    With p_ChrAttrs
      Sleep p_BwData.Delay
      Refresh
      If p_BwData.CriStop And .Cri >= p_BwData.CriPoints Then
        p_bRun = False
        Exit Do
      End If
      If .Str + .Dex + .Con + .Int + .Wis + .Cha >= p_BwData.Points Then
        p_bRun = False
        Exit Do
      End If
    End With
    SetCursorPos p_BwData.scrX, p_BwData.scrY
    mouse_event MEF_LEFTDOWN Or MEF_LEFTUP, p_BwData.scrX, p_BwData.scrY, 0, 0
  Loop
End Function

Public Sub SetDefalueValue(gmType As Byte)
  If gmType < GAME_NUM Then
    p_BwData.GameType = gmType
    With p_BwItem
      .WindowName = ""
      .BeginAddress = gmBeginAddr
      .EndAddress = gmEndAddr
      If Not .CustomKey Then
        .MemoryKeys = GetMemoryKeys(gmType)
        .Offset = GetOffSet(gmType)
      End If
    End With
    p_BwData.KeyStart = keyNumStar
    p_BwData.KeyStop = KeyNumSlash
  End If
End Sub

Public Sub SetMemoryKeys(Keys As String)
  Dim i As Integer
  Dim S() As String
  Dim byKeys() As Byte
  
  On Error GoTo setkeysHandler
  
  S = Split(Keys, ",")
  If UBound(S) <> 6 Then
    Exit Sub
  End If

  ReDim byKeys(0 To UBound(S))
  byKeys(Str) = CByte(CInt(S(0)))
  byKeys(Cri) = CByte(CInt(S(1)))
  byKeys(Dex) = CByte(CInt(S(2)))
  byKeys(Con) = CByte(CInt(S(3)))
  byKeys(Inl) = CByte(CInt(S(4)))
  byKeys(Wis) = CByte(CInt(S(5)))
  byKeys(Cha) = CByte(CInt(S(6)))
  p_BwItem.MemoryKeys = byKeys
  p_BwItem.Offset = 0
  p_BwItem.CustomKey = True

setkeysHandler:
End Sub

Public Function GetMemoryKeys(gmType As Byte) As Byte()
  Dim i As Integer
  Dim S() As String
  Dim byKeys() As Byte
  
  On Error GoTo getkeysHandler
  
  If gmType < BG1 Or gmType > HOW Then
    GetMemoryKeys = Null
  Else
    Select Case gmType
      Case BG1, TSC:
        S = Split(gmKeyBg1, ",")
      Case SOA, TOB:
        S = Split(gmKeySoa, ",")
      Case IWD, HOW:
        S = Split(gmKeyIwd, ",")
    End Select
    ReDim byKeys(0 To UBound(S))
    For i = 0 To UBound(S)
      byKeys(i) = CByte(S(i))
    Next
    GetMemoryKeys = byKeys
  End If
  Exit Function
  
getkeysHandler:
End Function

Public Function GetOffSet(gmType As Byte) As Integer
  Dim S() As String
  
  On Error GoTo getoffsetHandler
  
  If gmType < BG1 Or gmType > HOW Then
    GetOffSet = 0
  Else
    S = Split(gmOffset, ",")
    GetOffSet = S(gmType)
  End If
  Exit Function
  
getoffsetHandler:
End Function

Public Function AddNotifyIcon(ByVal hIcon As Long, ByVal hWnd As Long, strTip As String, dwMessage As Long) As Long
  If p_Nid.Status = 0 Then
    With p_Nid.Nid
      .cbSize = Len(p_Nid.Nid)
      .hIcon = hIcon
      .hWnd = hWnd
      If Len(strTip) > 63 Then
        .szTip = Left(strTip, 63) & Chr(0)
      Else
        .szTip = strTip & Chr(0)
      End If
      If dwMessage <> 0 Then
        .uCallbackMessage = dwMessage
      Else
        .uCallbackMessage = WM_MOUSEMOVE
      End If
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    End With
    p_Nid.Status = Shell_NotifyIcon(NIM_ADD, p_Nid.Nid)
  End If
  AddNotifyIcon = p_Nid.Status
End Function

Public Function RemoveNotifyIcon() As Long
  If p_Nid.Status <> 0 Then
    RemoveNotifyIcon = Shell_NotifyIcon(NIM_DELETE, p_Nid.Nid)
    If RemoveNotifyIcon Then
      p_Nid.Status = 0
    End If
  End If
  
End Function

Public Function DoSearch() As Boolean
  Dim lpBuff(0 To CHRATTRS_COUNT - 1) As Byte
  Dim lpNOBR As Long
  
  DoSearch = False
  On Error GoTo sechErrHandler
  If p_hProcess <> 0 Then
    CloseHandle p_hProcess
  End If
  p_hProcess = GetProcess(p_BwItem.ClassName, p_BwItem.WindowName)
  If p_hProcess <> 0 Then
    p_BwData.Address = Search(p_hProcess, p_BwItem.BeginAddress, p_BwItem.EndAddress, p_BwItem.MemoryKeys)
    If p_BwData.Address <> 0 Then
      p_BwData.Address = p_BwData.Address + p_BwItem.Offset
      If ReadProcessMemory(p_hProcess, p_BwData.Address, lpBuff(0), CHRATTRS_COUNT, lpNOBR) Then
        With p_ChrAttrs
          .Cha = lpBuff(Cha)
          .Con = lpBuff(Con)
          .Cri = lpBuff(Cri)
          .Dex = lpBuff(Dex)
          .Int = lpBuff(Inl)
          .Str = lpBuff(Str)
          .Wis = lpBuff(Wis)
        End With
      Else
        With p_ChrAttrs
          .Cha = 0
          .Con = 0
          .Cri = 0
          .Dex = 0
          .Int = 0
          .Str = 0
          .Wis = 0
        End With
      End If
      DoSearch = True
    End If
  End If
sechErrHandler:
End Function

Public Sub Refresh()
  Dim lpBuff(0 To CHRATTRS_COUNT - 1) As Byte
  Dim lpNOBR As Long
  
  On Error GoTo rfhErrHandler
  If p_hProcess = 0 Then
    p_hProcess = GetProcess(p_BwItem.ClassName, p_BwItem.WindowName)
  End If
  If p_hProcess <> 0 Then
    If ReadProcessMemory(p_hProcess, p_BwData.Address, lpBuff(0), CHRATTRS_COUNT, lpNOBR) Then
      With p_ChrAttrs
        .Cha = lpBuff(Cha)
        .Con = lpBuff(Con)
        .Cri = lpBuff(Cri)
        .Dex = lpBuff(Dex)
        .Int = lpBuff(Inl)
        .Str = lpBuff(Str)
        .Wis = lpBuff(Wis)
      End With
    End If
  End If
  
rfhErrHandler:
End Sub

Public Property Get ClassName() As String
  ClassName = p_BwItem.ClassName
End Property

Public Property Let ClassName(lpClassName As String)
  If lpClassName <> "" Then
    p_BwItem.ClassName = lpClassName
  End If
End Property

Public Property Get WindowName() As String
  WindowName = p_BwItem.WindowName
End Property

Public Property Let WindowName(ByVal lpWindowName As String)
  p_BwItem.WindowName = lpWindowName
End Property

Public Property Get BeginAddress() As Long
  BeginAddress = p_BwItem.BeginAddress
End Property

Public Property Let BeginAddress(ByVal lAddr As Long)
  If lAddr > 0 Then
    p_BwItem.BeginAddress = lAddr
  End If
End Property

Public Property Get EndAddress() As Long
  EndAddress = p_BwItem.EndAddress
End Property

Public Property Let EndAddress(ByVal lAddr As Long)
  If lAddr > 0 Then
    p_BwItem.EndAddress = lAddr
  End If
End Property

Public Property Get CustomKey() As Integer
  CustomKey = p_BwItem.CustomKey
End Property

Public Property Get Offset() As Integer
  Offset = p_BwItem.Offset
End Property

Public Property Let Offset(ByVal iOffset As Integer)
  If iOffset > -1 Then
    p_BwItem.Offset = iOffset
  End If
End Property

Public Property Get GameType() As Byte
  GameType = p_BwItem.GameType
End Property

Public Property Let GameType(ByVal gmType As Byte)
  If gmType < GAME_NUM Then
    p_BwItem.GameType = gmType
    p_BwData.GameType = gmType
  End If
End Property

Public Property Get MemoryKeys() As Byte()
  MemoryKeys = p_BwItem.MemoryKeys
End Property

Public Property Let MemoryKeys(ByRef Keys() As Byte)
  If Not IsNull(Keys) Then
    If UBound(Keys) - LBound(Keys) > 0 Then
      p_BwItem.MemoryKeys = Keys
    End If
  End If
End Property

Public Property Get Points() As Byte
  Points = p_BwData.Points
End Property

Public Property Let Points(ByVal byPoints As Byte)
  If byPoints < MAX_POINTS + 1 Then
    p_BwData.Points = byPoints
  End If
End Property

Public Property Get CriPoints() As Byte
  CriPoints = p_BwData.CriPoints
End Property

Public Property Let CriPoints(ByVal byCriPoints As Byte)
  If byCriPoints < MAX_CRIPOINTS + 1 Then
    p_BwData.CriPoints = byCriPoints
  End If
End Property

Public Property Get Delay() As Long
  Delay = p_BwData.Delay
End Property

Public Property Let Delay(ByVal iDelay As Long)
  If iDelay > 0 Then
    p_BwData.Delay = iDelay
  End If
End Property

Public Property Get CriStop() As Boolean
  CriStop = CBool(p_BwData.CriStop)
End Property

Public Property Let CriStop(ByVal bCriStop As Boolean)
  p_BwData.CriStop = CByte(bCriStop)
End Property

Public Property Get KeyStart() As Integer
  KeyStart = p_BwData.KeyStart
End Property

Public Property Let KeyStart(ByVal iKey As Integer)
  If iKey > -1 Then
    p_BwData.KeyStart = iKey
  End If
End Property

Public Property Get KeyStop() As Integer
  KeyStop = p_BwData.KeyStop
End Property

Public Property Let KeyStop(ByVal iKey As Integer)
  If iKey > -1 Then
    p_BwData.KeyStop = iKey
  End If
End Property

Public Property Get Strength() As Byte
  Strength = p_ChrAttrs.Str
End Property

Public Property Get Dexterity() As Byte
  Dexterity = p_ChrAttrs.Dex
End Property

Public Property Get Constitution() As Byte
  Constitution = p_ChrAttrs.Con
End Property

Public Property Get Intelligence() As Byte
  Intelligence = p_ChrAttrs.Int
End Property

Public Property Get Wisdom() As Byte
  Wisdom = p_ChrAttrs.Wis
End Property

Public Property Get Charisma() As Byte
  Charisma = p_ChrAttrs.Cha
End Property

Public Property Get Crital() As Byte
  Crital = p_ChrAttrs.Cri
End Property

Public Property Get Captions() As String()
  Captions = Split(gmCaption, ",")
End Property

Public Property Get Titles() As String()
  Titles = Split(gmTitle, ",")
End Property

Public Property Get MaxPoints() As Integer
  MaxPoints = MAX_POINTS
End Property

Public Property Get MaxCriPoints() As Integer
  MaxCriPoints = MAX_CRIPOINTS
End Property

Public Property Get scrX() As Integer
  scrX = p_BwData.scrX
End Property

Public Property Let scrX(ByVal iValue As Integer)
  If iValue > -1 And iValue < Screen.Width / Screen.TwipsPerPixelX Then
    p_BwData.scrX = iValue
  End If
End Property

Public Property Get scrY() As Integer
  scrY = p_BwData.scrY
End Property

Public Property Let scrY(ByVal iValue As Integer)
  If iValue > -1 And iValue < Screen.Height / Screen.TwipsPerPixelY Then
    p_BwData.scrY = iValue
  End If
End Property


Public Property Get EnableCheck() As Boolean
  EnableCheck = p_bRun
End Property

Public Property Let EnableCheck(ByVal bValue As Boolean)
  If bValue Then
    If Not p_bRun Then
      p_bRun = True
      CheckAttrs
    End If
  Else
    p_bRun = False
  End If
End Property


Public Sub SetTimer(tmObj As Timer)
  If IsObject(tmObj) Then
    Set p_Timer = tmObj
    p_Timer.Enabled = False
  End If
End Sub

Public Function GetTimer() As Timer
  GetTimer = p_Timer
End Function

Public Sub StartCheck()
  On Error GoTo errHandler
  
  If IsObject(p_Timer) Then
    p_Timer.Interval = p_BwData.Delay
    p_Timer.Enabled = True
  End If
errHandler:
End Sub

Public Sub StopCheck()
  p_Timer.Enabled = False
End Sub

Private Sub p_Timer_Timer()
 
  With p_ChrAttrs
    Refresh
    If p_BwData.CriStop And .Cri >= p_BwData.CriPoints Then
        p_Timer.Enabled = False
        Exit Sub
    End If
    If .Str + .Dex + .Con + .Int + .Wis + .Cha >= p_BwData.Points Then
      p_Timer.Enabled = False
      Exit Sub
    End If
  End With
  SetCursorPos p_BwData.scrX, p_BwData.scrY
  mouse_event MEF_LEFTDOWN Or MEF_LEFTUP, p_BwData.scrX, p_BwData.scrY, 0, 0
End Sub
