VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TbgWoods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum GameType
  BG1 = 0
  TSC = 1
  SOA = 2
  TOB = 3
  IWD = 4
  HOW = 5
End Enum

Enum AttrType
  CHRSTR = 0
  CHRDEX = 4
  CHRCON = 5
  CHRINT = 2
  CHRWIS = 3
  CHRCHA = 6
  CHRCRI = 1
End Enum

Private FGTStr(0 To 5) As String
Private FGNStr(0 To 5) As String
Private FChrAttrs As CChrAttrs
Private FBgData As CBgData

Private Sub SetValues()
Dim count As Integer

  FGTStr(BG1) = "BG1"
  FGTStr(TSC) = "TSC"
  FGTStr(SOA) = "SOA"
  FGTStr(TOB) = "TOB"
  FGTStr(IWD) = "IWD"
  FGTStr(HOW) = "HOW"
  FGNStr(BG1) = "博得之门"
  FGNStr(TSC) = "博得之门：剑湾传奇"
  FGNStr(SOA) = "博得之门二：安姆的阴影"
  FGNStr(TOB) = "博得之门二：巴尔王座"
  FGNStr(IWD) = "冰风溪谷"
  FGNStr(HOW) = "冰风溪谷：冬之心"
  With FBgData
    .Address = 0
    .Points = 0
    .CriPoints = 0
    .CriStop = 0
    .Delay = 0
    .GameType = 3
  End With
  With FChrAttrs
    .Str = 0
    .Dex = 0
    .Con = 0
    .Int = 0
    .Wis = 0
    .Cha = 0
    .Cri = 0
  End With
End Sub

Public Function GetName(gtValue As GameType) As String
  GetName = FGNStr(gtValue)
End Function

Public Function TypeToStr(gtValue As GameType) As String
  TypeToStr = FGTStr(gtValue)
End Function

Private Sub Class_Initialize()
  Call SetValues
End Sub

Public Property Get MaxPoints() As Byte
  MaxPoints = 108
End Property

Public Property Get MaxCriPoints() As Byte
  MaxCriPoints = 100
End Property

Public Property Get Strength() As Byte
  Strength = FChrAttrs.Str
End Property
Public Property Get Dexterity() As Byte
  Dexterity = FChrAttrs.Dex
End Property

Public Property Get Constitution() As Byte
  Constitution = FChrAttrs.Con
End Property

Public Property Get Intelligence() As Byte
  Intelligence = FChrAttrs.Int
End Property

Public Property Get Wisdom() As Byte
  Wisdom = FChrAttrs.Wis
End Property

Public Property Get Charisma() As Byte
  Charisma = FChrAttrs.Cha
End Property

Public Property Get Critical() As Byte
  Critical = FChrAttrs.Cri
End Property

Public Property Let GameType(Val As Byte)
  FBgData.GameType = Val
End Property

Public Property Let Points(Val As Byte)
  FBgData.Points = Val
End Property

Public Property Let CriPoints(Val As Byte)
  FBgData.CriPoints = Val
End Property

Public Property Let CriStop(Val As Byte)
  FBgData.CriStop = Val
End Property

Public Property Let Delay(Val As Long)
  FBgData.Delay = Val
End Property

Public Property Get GameType() As Byte
  GameType = FBgData.GameType
End Property

Public Property Get Points() As Byte
  Points = FBgData.Points
End Property

Public Property Get CriPoints() As Byte
  CriPoints = FBgData.CriPoints
End Property

Public Property Get CriStop() As Byte
  CriStop = FBgData.CriStop
End Property

Public Property Get Delay() As Long
  Delay = FBgData.Delay
End Property

Public Property Get Address() As Long
  Address = FBgData.Address
End Property

Public Sub SearchChrAttrs(Val As String)
  Dim S() As String
  Dim Attrs As CChrAttrs
  
  Attrs.Str = 0
  S = Split(Val, ",")
  If UBound(S) = 6 Then
    Attrs.Str = CByte(CInt(S(0)))
    Attrs.Cri = CByte(CInt(S(1)))
    Attrs.Dex = CByte(CInt(S(2)))
    Attrs.Con = CByte(CInt(S(3)))
    Attrs.Int = CByte(CInt(S(4)))
    Attrs.Wis = CByte(CInt(S(5)))
    Attrs.Cha = CByte(CInt(S(6)))
  End If
  SetData FBgData
  FChrAttrs = Search(Attrs)
End Sub

Public Sub SetBgData()
  SetData FBgData
End Sub
