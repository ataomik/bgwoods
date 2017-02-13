Attribute VB_Name = "mdlApi"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNid As NotifyIconData) As Boolean
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Const GWL_WNDPROC = -4

Public Const WM_GETMINMAXINFO = &H24
Public Const WM_SIZE = &H5
Public Const WM_MOUSEMOVE = &H200
Public Const SIZE_MAXHIDE = 4
Public Const SIZE_MAXIMIZED = 2
Public Const SIZE_MAXSHOW = 3
Public Const SIZE_MINIMIZED = 1
Public Const SIZE_RESTORED = 0

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const SW_SHOWNORMAL = 1

Public Type NotifyIconData
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Type CPoint
  x As Long
  y As Long
End Type

Public Type MINMAXINFO
  ptReserved As CPoint
  ptMaxSize As CPoint
  ptMaxPosition As CPoint
  ptMinTrackSize As CPoint
  ptMaxTrackSize As CPoint
End Type

Public Type CMinMax
  MinWidth As Long
  MinHeight As Long
  MaxWidth As Long
  MaxHeight As Long
End Type

Public oldWndProc As Long
Public defMinMax As CMinMax

Public Function SubClass_WndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

  If Msg = WM_GETMINMAXINFO Then
  Dim MinMaxSize As MINMAXINFO

    CopyMemory MinMaxSize, ByVal lp, Len(MinMaxSize)
  
    MinMaxSize.ptMinTrackSize.x = defMinMax.MinWidth \ Screen.TwipsPerPixelX
    MinMaxSize.ptMinTrackSize.y = defMinMax.MinHeight \ Screen.TwipsPerPixelY
    MinMaxSize.ptMaxTrackSize.x = defMinMax.MaxWidth \ Screen.TwipsPerPixelX
    MinMaxSize.ptMaxTrackSize.y = defMinMax.MaxHeight \ Screen.TwipsPerPixelY
        
    CopyMemory ByVal lp, MinMaxSize, Len(MinMaxSize)
    SubClass_WndMessage = 1
    Exit Function
  End If
  
  SubClass_WndMessage = CallWindowProc(oldWndProc, hWnd, Msg, wp, lp)
End Function



