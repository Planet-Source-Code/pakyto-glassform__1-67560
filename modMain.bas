Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean
Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public ProcOld As Long
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Const WM_GETMINMAXINFO As Long = &H24
Public Const GWL_WNDPROC = (-4)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
Public Const WS_MINIMIZEBOX = &H20000
Private Const GW_HWNDNEXT = 2
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1


Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case iMsg
Case WM_GETMINMAXINFO
Dim udtMINMAXINFO As MINMAXINFO
CopyMemory udtMINMAXINFO, ByVal lParam, 40&
With udtMINMAXINFO
.ptMinTrackSize.x = 230
.ptMinTrackSize.y = 180
End With
CopyMemory ByVal lParam, udtMINMAXINFO, 40&
WindowProc = False
Exit Function
End Select
WindowProc = CallWindowProc(ProcOld, hWnd, iMsg, wParam, lParam)
End Function


Public Sub RedondearEsquinaForm(pForm As Form, Radio As Long)

'Variables varias
Dim Ret As Long
Dim l As Long
Dim llWidth As Long
Dim llHeight As Long

'Obtenemos el ancho y alto de la region del Form
llWidth = pForm.Width / Screen.TwipsPerPixelX
llHeight = pForm.Height / Screen.TwipsPerPixelY

'Le pasamos el ancho alto del formualrio y el valor de redondeo es decir el radio a la función Api
Ret = CreateRoundRectRgn(0, 0, llWidth, llHeight, Radio, Radio)

'Se lo aplicamos
l = SetWindowRgn(pForm.hWnd, Ret, True)
End Sub

Public Sub SetTrans(oForm As Form, Optional bytAlpha As Byte = 255, Optional lColor As Long = 0)
    Dim lStyle As Long
    lStyle = GetWindowLong(oForm.hWnd, GWL_EXSTYLE)
    If Not (lStyle And WS_EX_LAYERED) = WS_EX_LAYERED Then _
        SetWindowLong oForm.hWnd, GWL_EXSTYLE, lStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes oForm.hWnd, lColor, bytAlpha, LWA_COLORKEY Or LWA_ALPHA
End Sub

