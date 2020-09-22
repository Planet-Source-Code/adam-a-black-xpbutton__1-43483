Attribute VB_Name = "hTimerBas"
Option Explicit

'**API Calls
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'** Define Constants
Public Const WS_MINIMIZE = &H20000000
Public Const WS_OVERLAPPED = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const WM_USER As Long = &H400
Public Const WM_TIMER_FIRED As Long = (WM_USER + &H1001)

Private Const hAuth As Long = &H34989812

'** Define Objects/Variables
Private ClassCollection As New Collection
Private mlOriginalWinProc As Long

Sub TimerProc(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
Attribute TimerProc.VB_Description = "Special CallBack Procedure - DO NOT USE!"
Attribute TimerProc.VB_MemberFlags = "40"
    Call PostMessageA(dwUser, WM_TIMER_FIRED, 0, dwUser)
End Sub

Public Function hRegister(hTmr As hTimerCls)
    Dim dwHandle As Long
    
    dwHandle = CreateWindowEx(0, "STATIC", "hTimerWindow", WS_OVERLAPPED Or WS_MINIMIZE, 0, 0, 100, 100, 0, 0, 0, 0)
    
    If Not dwHandle = 0 Then
        Call fStartMessageCapture(dwHandle)
        ClassCollection.Add hTmr, CStr(dwHandle)
    End If
    
    hRegister = dwHandle
End Function

Public Function hUnregister(dwHandle As Long) As Long
    On Error Resume Next
    
    Call fEndMessageCapture(dwHandle)
    DestroyWindow dwHandle
    ClassCollection.Remove CStr(dwHandle)
End Function

Public Sub fStartMessageCapture(ByVal hwnd As Long)
    Dim lWinProc As Long
    
    mlOriginalWinProc = 0
    
    lWinProc = GetWindowLong(hwnd, GWL_WNDPROC)
    
    If lWinProc <> 0 Then
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf fLocWinProc)
    End If
    
    mlOriginalWinProc = lWinProc
End Sub

Public Sub fEndMessageCapture(ByVal hwnd As Long)
    Call SetWindowLong(hwnd, GWL_WNDPROC, mlOriginalWinProc)
End Sub

Public Function fLocWinProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim CallingClass As hTimerCls
    
    Select Case uMsg
        Case WM_TIMER_FIRED:
            Set CallingClass = ClassCollection(CStr(lParam))
            CallingClass.Timer_Event (hAuth)
            Set CallingClass = Nothing
            fLocWinProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
        Case Else:
            fLocWinProc = CallWindowProc(mlOriginalWinProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function

