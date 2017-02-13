Attribute VB_Name = "Winapi"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNid As NotifyIconData) As Boolean
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'------------------------------------------------------------------------------------------------ variable declare

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
Public Const MIN_WIDTH = 6500
Public Const MIN_HEIGHT = 3750
Public Const MAX_WIDTH = 6500
Public Const MAX_HEIGHT = 3750

Public Type NotifyIconData
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public OldWndProc As Long
Public Result As Long

Public Function SubClass_WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

    If Msg = WM_GETMINMAXINFO Then

        Dim MinMaxSize As MINMAXINFO
        
        CopyMemory MinMaxSize, ByVal lp, Len(MinMaxSize)
  
        MinMaxSize.ptMinTrackSize.X = MIN_WIDTH \ Screen.TwipsPerPixelX
        MinMaxSize.ptMinTrackSize.Y = MIN_HEIGHT \ Screen.TwipsPerPixelY
        MinMaxSize.ptMaxTrackSize.X = MAX_WIDTH \ Screen.TwipsPerPixelX
        MinMaxSize.ptMaxTrackSize.Y = MAX_HEIGHT \ Screen.TwipsPerPixelY
        
        CopyMemory ByVal lp, MinMaxSize, Len(MinMaxSize)
        SubClass_WndMessage = 1
        Exit Function
    End If
    
    SubClass_WndMessage = CallWindowProc(OldWndProc, hwnd, Msg, wp, lp)
        
End Function



