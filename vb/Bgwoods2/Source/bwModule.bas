Attribute VB_Name = "bwModule"
Type CPoint
  X As Long
  Y As Long
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

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Public Const GWL_WNDPROC = -4

Public Const WM_GETMINMAXINFO = &H24
Public Const WM_SIZE = &H5
Public Const SIZE_MAXHIDE = 4
Public Const SIZE_MAXIMIZED = 2
Public Const SIZE_MAXSHOW = 3
Public Const SIZE_MINIMIZED = 1
Public Const SIZE_RESTORED = 0
Public Const SW_NORMAL = 1
Public Const WM_MOUSEMOVE = &H200
  
Public oldWndProc As Long
Public defMinMax As CMinMax

Public Function SubClass_WndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

  If Msg = WM_GETMINMAXINFO Then
  Dim MinMaxSize As MINMAXINFO
    CopyMemory MinMaxSize, ByVal lp, Len(MinMaxSize)
  
    MinMaxSize.ptMinTrackSize.X = defMinMax.MinWidth \ Screen.TwipsPerPixelX
    MinMaxSize.ptMinTrackSize.Y = defMinMax.MinHeight \ Screen.TwipsPerPixelY
    MinMaxSize.ptMaxTrackSize.X = defMinMax.MaxWidth \ Screen.TwipsPerPixelX
    MinMaxSize.ptMaxTrackSize.Y = defMinMax.MaxHeight \ Screen.TwipsPerPixelY
        
    CopyMemory ByVal lp, MinMaxSize, Len(MinMaxSize)
    SubClass_WndMessage = 1
    Exit Function
  End If
    
  SubClass_WndMessage = CallWindowProc(oldWndProc, hWnd, Msg, wp, lp)
End Function
