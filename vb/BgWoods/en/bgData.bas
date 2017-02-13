Attribute VB_Name = "bgData"
Public Declare Function EnableHotKeyHook Lib "Bgrun.dll" () As Long
Public Declare Function DisableHotKeyHook Lib "Bgrun.dll" () As Long
Public Declare Function SetData Lib "Bgrun.dll" (ByRef lpData As CBgData) As CBgData
Public Declare Function Search Lib "Bgrun.dll" (ByRef Val As CChrAttrs) As CChrAttrs

Public Type CBgData
  GameType As Byte
  Points As Byte
  CriPoints As Byte
  CriStop As Byte
  Delay As Long
  Address As Long
End Type

Public Type CChrAttrs
  Str As Byte
  Dex As Byte
  Con As Byte
  Int As Byte
  Wis As Byte
  Cha As Byte
  Cri As Byte
End Type
