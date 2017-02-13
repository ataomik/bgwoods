Attribute VB_Name = "CBwMod"
Option Explicit

Private Type CPoint
  x As Long
  y As Long
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotkey Lib "user32" Alias "RegisterHotKey" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnRegisterHotkey Lib "user32" Alias "UnregisterHotKey" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As CPoint) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_KEYDOWN = &H100
Private Const WM_HOTKEY = &H312
Private Const ID_KEYSTART = "911"
Private Const ID_KEYSTOP = "8012"

Private BwRunlib As CBwLib
Private oldWndProc As Long

Private Function SubClass_BwWndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim ptScreen As CPoint

  If Msg = WM_HOTKEY Then
    If wParam = ID_KEYSTART Then
      If Not BwRunlib.EnableCheck Then
        GetCursorPos ptScreen
        BwRunlib.scrX = ptScreen.x
        BwRunlib.scrY = ptScreen.y
        BwRunlib.StartCheck
      End If
    Else
      If wParam = ID_KEYSTOP Then
        BwRunlib.StopCheck
      End If
    End If
  End If
    
  SubClass_BwWndMessage = CallWindowProc(oldWndProc, hWnd, Msg, wParam, lParam)
End Function

Public Function EnableHotkey(ByVal hWnd As Long, Bwlib As CBwLib) As Long
  EnableHotkey = 0
  If Not (Bwlib Is Nothing) Then
    Set BwRunlib = Bwlib
    If hWnd <> 0 Then
      oldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubClass_BwWndMessage)
      If oldWndProc <> 0 Then
        RegisterHotkey hWnd, ID_KEYSTART, 0, BwRunlib.KeyStart
        RegisterHotkey hWnd, ID_KEYSTOP, 0, BwRunlib.KeyStop
        EnableHotkey = oldWndProc
      End If
    End If
  End If
End Function

Public Function DisableHotkey(hWnd As Long) As Long
  
  If oldWndProc <> 0 And (Not (BwRunlib Is Nothing)) Then
    If hWnd <> 0 Then
      UnRegisterHotkey hWnd, ID_KEYSTART
      UnRegisterHotkey hWnd, ID_KEYSTOP
      SetWindowLong hWnd, GWL_WNDPROC, oldWndProc
    End If
  End If
End Function
