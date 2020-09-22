VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "hTimerCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'** Define Constants
Private Const TIME_CALLBACK_FUNCTION = &H0
Private Const TIME_PERIODIC = 1
Private Const hAuth As Long = &H34989812

'** Define Variables
Private lngInterval As Long
Private bolEnabled As Boolean
Private tHandle As Long
'** Define Events
Public Event Timer()

'** Declare DLL Function Calls
Private Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
Private Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long

Private Sub Class_Initialize()
    '** Initilize default values
    bolEnabled = False
    lngInterval = 1000
End Sub

Private Sub Class_Terminate()
    '** Clean Up
    Enabled = False
End Sub

'** Interval Property
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer's Timer event."
    Interval = lngInterval
End Property

Public Property Let Interval(ByVal vNewValue As Long)
    If (bolEnabled = True) And (vNewValue <> lngInterval) Then
        lngInterval = vNewValue
        Enabled = False
        Enabled = True
    End If
    
    lngInterval = vNewValue
End Property

'** Enabled Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = bolEnabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Static tmrID As Long
    
    '** If supplied Value is the same then don't continue
    If vNewValue = bolEnabled Then Exit Property
    
    '** Assign to Variable
    bolEnabled = vNewValue

    '** Enable or Disable timer
    If vNewValue Then
        '** Creater Timer Window
        If tHandle = 0 Then tHandle = hRegister(Me)
        '** Create Timer
        tmrID = timeSetEvent(lngInterval, 1, AddressOf TimerProc, tHandle, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    Else
        '** Destroy Active Timer
        timeKillEvent tmrID
        tmrID = 0
            
        '** Destroy Timer Window
        hUnregister tHandle
        tHandle = 0
    End If
    
End Property

Public Sub Timer_Event(hCode As Long)
Attribute Timer_Event.VB_Description = "Special CallBack Procedure - DO NOT USE!"
Attribute Timer_Event.VB_MemberFlags = "40"
    If hCode = hAuth Then RaiseEvent Timer
End Sub
