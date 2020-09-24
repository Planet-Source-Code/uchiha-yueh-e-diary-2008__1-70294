Attribute VB_Name = "Modular"
'For program on top..
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
 
'For sounds..
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_SYNC = &H0         '  play synchronously (default)
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'For dragging.........
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const WM_NCLBUTTONDOWN = &HA1

Public Const HTCAPTION = 2

Public JScript As String

Public NewText As String


Public Sub DragIt(ByVal lngHwnd As Long)

Dim lngReturn As Long

    lngReturn = ReleaseCapture()
    lngReturn = SendMessage(lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, CLng(0))
    
End Sub

'// Resets the windows position in the ZOrder
Public Sub MakeNormal(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

'// Places the window at the top of the ZOrder
'// and keeps it there
Public Sub MakeTopMost(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
