VERSION 5.00
Begin VB.UserControl XPButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   DefaultCancel   =   -1  'True
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   73
   ToolboxBitmap   =   "XPButton.ctx":0000
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***----API declarations...
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'***----/API declarations...

'***----API constants-------------------
Private Const RDW_INVALIDATE = &H1
Private Const BITSPIXEL = 12
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const PS_SOLID = 0
'***----/API constants------------------

'***----API types-----------------
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type BITMAPINFO
    Header  As BITMAPINFOHEADER
    Bits()  As Byte
End Type

Private Type POINTAPI
    x   As Long
    y   As Long
End Type
'***----/API types----------------

'***--custom types/enums-------------
Private Type RGBType
    Red     As Integer
    Green   As Integer
    Blue    As Integer
End Type

'Different border styles
Enum BorderStyleEnum
    Default_XP = 1
    Custom_1 = 2
    Custom_2 = 3
    Custom_3 = 4
    Custom_4 = 5
    Custom_5 = 6
    Custom_6 = 7
    Custom_7 = 8
End Enum

Enum PictureAlignmentEnum
    Left = 1
    Right = 2
    Top = 3
    Bottom = 4
    Center = 5
End Enum

Enum BackColourEnum
    Light = 1
    Medium = 2
    Dark = 3
End Enum

Enum BackStyleEnum
    Solid = 1
    Gradient = 2
End Enum
'***--/custom types/enums------------

'-----------STD COLOURS--------------'
Const TopBGrad      As Long = 16514300  'top background gradient
Const BotBGrad      As Long = 15397104  'bottom background gradient
Const EndBGrad      As Long = 15133676  'colour below gradient
Const OutBorder     As Long = 7617536   'outside border colour
'---------END STD COLOURS------------'

'------------DRAW FOCUS--------------'
'--------DEFAULT XP SCHEME-----------'
Const TopFirstF1    As Long = 16771022
Const TopSecondF1   As Long = 16176316
Const BotFirstF1    As Long = 14986633
Const BotSecondF1   As Long = 15630953
Const OutPixelF     As Long = 11048314
Const InPixelF      As Long = 8736039
'----------CUSTOM SCHEME 1------------'
Const TopFirstF2    As Long = 7405298
Const TopSecondF2   As Long = 60633
Const BotFirstF2    As Long = 574131
Const BotSecondF2   As Long = 565657
'----------CUSTOM SCHEME 2------------'
Const TopFirstF3    As Long = 13485251
Const TopSecondF3   As Long = 13545902
Const BotFirstF3    As Long = 12545895
Const BotSecondF3   As Long = 12017751
'----------CUSTOM SCHEME 3------------'
Const TopFirstF4    As Long = 10091514
Const TopSecondF4   As Long = 7728106
Const BotFirstF4    As Long = 6142906
Const BotSecondF4   As Long = 5018257
'----------CUSTOM SCHEME 4------------'
Const TopFirstF5    As Long = 16767465
Const TopSecondF5   As Long = 16762334
Const BotFirstF5    As Long = 16744886
Const BotSecondF5   As Long = 16604577
'----------CUSTOM SCHEME 5------------'
Const TopFirstF6    As Long = 16761251
Const TopSecondF6   As Long = 16757131
Const BotFirstF6    As Long = 13583877
Const BotSecondF6   As Long = 10040837
'----------CUSTOM SCHEME 6------------'
Const TopFirstF7    As Long = 367547
Const TopSecondF7   As Long = 441056
Const BotFirstF7    As Long = 6478560
Const BotSecondF7   As Long = 11725558
'----------CUSTOM SCHEME 7------------'
Const TopFirstF8    As Long = 15630953
Const TopSecondF8   As Long = 14986633
Const BotFirstF8    As Long = 16176316
Const BotSecondF8   As Long = 16771022
'----------END DRAW FOCUS------------'

'-----------DRAW MOUSEIN-------------'
'--------DEFAULT XP SCHEME-----------'
Const TopFirstM1    As Long = 13627647
Const TopSecondM1   As Long = 9033981
Const BotFirstM1    As Long = 3191800
Const BotSecondM1   As Long = 38885
Const OutPixelM     As Long = 11048314
Const InPixelM  As Long = 10591885
'----------CUSTOM SCHEME 1------------'
Const TopFirstM2    As Long = 13358847
Const TopSecondM2   As Long = 11123199
Const BotFirstM2    As Long = 8492539
Const BotSecondM2   As Long = 6255863
'----------CUSTOM SCHEME 2------------'
Const TopFirstM3    As Long = 12243139
Const TopSecondM3   As Long = 10797232
Const BotFirstM3    As Long = 6073472
Const BotSecondM3   As Long = 4098143
'----------CUSTOM SCHEME 3------------'
Const TopFirstM4    As Long = 12648384
Const TopSecondM4   As Long = 5239631
Const BotFirstM4    As Long = 2739497
Const BotSecondM4   As Long = 2597671
'----------CUSTOM SCHEME 4------------'
Const TopFirstM5    As Long = 6029286
Const TopSecondM5   As Long = 63185
Const BotFirstM5    As Long = 51883
Const BotSecondM5   As Long = 176278
'----------CUSTOM SCHEME 5------------'
Const TopFirstM6    As Long = 391941
Const TopSecondM6   As Long = 252675
Const BotFirstM6    As Long = 300548
Const BotSecondM6   As Long = 285700
'----------CUSTOM SCHEME 6------------'
Const TopFirstM7    As Long = 5215055
Const TopSecondM7   As Long = 5618005
Const BotFirstM7    As Long = 9953687
Const BotSecondM7   As Long = 11922869
'----------CUSTOM SCHEME 7------------'
Const TopFirstM8    As Long = 38885
Const TopSecondM8   As Long = 3191544
Const BotFirstM8    As Long = 9033981
Const BotSecondM8   As Long = 13627647
'---------END DRAW MOUSEIN------------'

'-----------DRAW MOUSEDOWN------------'
Const TopFirstMD    As Long = 12700881
Const TopSecondMD   As Long = 13621468
Const TopThirdMD    As Long = 14542053
Const TopFourthMD   As Long = 14410468
Const BotFirstMD    As Long = 14936554
Const BotSecondMD   As Long = 15659506
Const LeftFirstMD   As Long = 13489881
Const LeftSecondMD  As Long = 14015967
Const OutPixelMD    As Long = 11048314
Const InPixelMD     As Long = 13353394
Const MidPixelMD    As Long = 9334086
'---------END DRAW MOUSEDOWN----------'

'-------------DRAW IDLE---------------'
Const TopFirstI     As Long = 16777215
Const TopSecondI    As Long = 16711422
Const BotFirstI     As Long = 14082018
Const BotSecondI    As Long = 12964054
Const OutPixelI     As Long = 11048314
Const InPixelI      As Long = 14470847
Const MidPixelI     As Long = 10648917
'-----------END DRAW IDLE-------------'

'-----------DRAW DISABLED-------------'
Const OutBorderD    As Long = 12240841
Const OutPixelD     As Long = 13096408
Const InPixelD      As Long = 14608874
Const MidPixelD     As Long = 13293272
Const ForeColourD   As Long = 9609633
'---------END DRAW DISABLED-----------'

'***--variables--------------------
Dim WithEvents hTimer As hTimerCls
Attribute hTimer.VB_VarHelpID = -1

Dim LowCol          As Boolean
Dim fInit           As Boolean
Dim HasFocus        As Boolean
Dim SpaceDown       As Boolean
Dim MouseDown       As Boolean
Dim MouseOver       As Boolean
Dim TopFirst        As Long
Dim TopSecond       As Long
Dim BotFirst        As Long
Dim BotSecond       As Long
Dim hBits           As BITMAPINFO
Dim mPos            As POINTAPI
Dim GradFoc()       As RGBType
Dim GradHot()       As RGBType
Dim GradBack()      As RGBType
'***--/variables-------------------

'Default Property Values:
Const m_def_PictureAlignment = 1
Const m_def_BackStyle               As Byte = 2
Const m_def_ButtonState             As Byte = 1
Const m_def_BackColour              As Byte = 2
Const m_def_BorderStyle             As Byte = 1
Const m_def_BackColourLow           As Long = 15793151
Const m_def_ForeColour              As Long = &H80000012
Const m_def_BackColourDown          As Long = 14279138
Const m_def_BackColourDownLow       As Long = 15793151
Const m_def_BackColour_light        As Long = 16382714
Const m_def_BackColour_med          As Long = 15857141
Const m_def_BackColour_dark         As Long = 15462640
Const m_def_DisabledBackColour      As Long = 15398133
Const m_def_DisableBackColourLow    As Long = 15793151
'Property Variables:
Dim m_BorderStyle       As Byte
Dim m_ButtonState       As Byte
Dim m_BackStyle         As Byte
Dim m_BackColour        As Byte
Dim m_PictureAlignment  As Byte
Dim m_AccessKey         As String * 1
Dim m_Caption           As String
Dim m_Picture           As String
Dim m_ForeColour        As OLE_COLOR
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseEnter()
Event MouseExit()
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Private Sub hTimer_Timer()
    Call GetCursorPos(mPos)
    
    'the mouse is not inside the usercontrol
    If Not WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd Then
        'the mouse was previously inside the usercontrol
        If m_ButtonState = 3 Then
            'the mouse is no longer over the usercontrol.
            MouseOver = False
            'the usercontrol previously had focus
            If HasFocus = True Then
                m_ButtonState = DrawButton(2)
            Else
                m_ButtonState = DrawButton(1)
            End If
            'raise the event
            RaiseEvent MouseExit
            'stop checking mouse position since the mouse has left.
            hTimer.Enabled = False
        End If
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If Extender.Default = True Then
        If Ambient.DisplayAsDefault = True And HasFocus = False Then
            Call UserControl_EnterFocus
        ElseIf Ambient.DisplayAsDefault = False And HasFocus = True Then
            Call UserControl_ExitFocus
        End If
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
    If HasFocus = False Then
        HasFocus = True
        Call GetCursorPos(mPos)
        
        If WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd Then
            m_ButtonState = DrawButton(3)
            MouseOver = True
            hTimer.Enabled = True
        Else
            m_ButtonState = DrawButton(2)
        End If
    End If
End Sub

Private Sub UserControl_ExitFocus()
    If (Extender.Default = True And Ambient.DisplayAsDefault = False) Or (Extender.Default = False) Then
        HasFocus = False
        If MouseDown = False And SpaceDown = False Then
            Call GetCursorPos(mPos)
            'the mouse is outside the usercontrol
            If Not (WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd) Then
                m_ButtonState = DrawButton(1)
            End If
        'the mouse button is clicked down
        ElseIf MouseDown = True Then
            Call GetCursorPos(mPos) 'get cursor position
            If WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd Then
                m_ButtonState = DrawButton(3)
                hTimer.Enabled = True
            Else
                m_ButtonState = DrawButton(1)
            End If
            MouseDown = False
            ReleaseCapture
        ElseIf SpaceDown = True Then
            SpaceDown = False
            m_ButtonState = DrawButton(1)
            ReleaseCapture
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Set hTimer = New hTimerCls
    hTimer.Interval = 2
        
    'returns true if colour depth is 8 or below. If true, no gradient is used
    'and non-dithering colours are used.
    LowCol = GetDeviceCaps(UserControl.hdc, BITSPIXEL) <= 8
    
    'set the default button state
    m_ButtonState = m_def_ButtonState
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If MouseDown = True Then Exit Sub
    'the user pushes space
    If KeyCode = vbKeySpace And UserControl.Enabled = True Then
        m_ButtonState = DrawButton(4)
        SetCapture UserControl.hwnd
        hTimer.Enabled = False
        SpaceDown = True
    'the user pushes return
    ElseIf KeyCode = vbKeyReturn And UserControl.Enabled = True Then
        m_ButtonState = DrawButton(1)
        RaiseEvent Click
        Call GetCursorPos(mPos)
        If Not WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd Then
            m_ButtonState = DrawButton(2)
            MouseOver = False
        Else
            m_ButtonState = DrawButton(3)
            MouseOver = True
            hTimer.Enabled = True
        End If
    'the user pushes a different key while the space bar is down.
    ElseIf SpaceDown = True Then
        Call GetCursorPos(mPos) 'get cursor position
        SpaceDown = False
        If WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd Then
            m_ButtonState = DrawButton(3)
            hTimer.Enabled = True
        Else
            m_ButtonState = DrawButton(2)
        End If
    ElseIf SpaceDown = False Then
        Select Case KeyCode
            Case vbKeyRight
                SendKeys "{TAB}"
            Case vbKeyLeft
                SendKeys "+{TAB}"
            Case vbKeyUp
                SendKeys "+{TAB}"
            Case vbKeyDown
                SendKeys "{TAB}"
        End Select
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    
    If SpaceDown = True And KeyCode = vbKeySpace Then
        ReleaseCapture
        Call GetCursorPos(mPos) 'get cursor position
        SpaceDown = False
        
        m_ButtonState = DrawButton(1)
        
        RaiseEvent Click
        
        Call GetCursorPos(mPos)
        If Not (WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd) Then
            m_ButtonState = DrawButton(2)
            MouseOver = False
        Else
            m_ButtonState = DrawButton(3)
            MouseOver = True
            hTimer.Enabled = True
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = Ambient.DisplayName
    m_BorderStyle = m_def_BorderStyle
    m_ForeColour = m_def_ForeColour
    m_BackColour = m_def_BackColour
    m_ButtonState = m_def_ButtonState
    m_BackStyle = m_def_BackStyle
    m_Picture = vbNullString
    m_PictureAlignment = m_def_PictureAlignment
    
    fInit = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    
    If SpaceDown = True Then Exit Sub
    
    If Button = vbLeftButton And UserControl.Enabled = True Then
        m_ButtonState = DrawButton(4)
        hTimer.Enabled = False
        MouseDown = True
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
    
    If SpaceDown = True Then Exit Sub
    
    'the mouse is over the usercontrol
    If x >= 0 And x <= UserControl.ScaleWidth And y >= 0 And y <= UserControl.ScaleHeight Then
        'the button hasn't been drawn mouseover and the mouse button isn't down,
        If (Not m_ButtonState = 3) And (MouseDown = False) Then
            m_ButtonState = DrawButton(3)
            MouseOver = True
            RaiseEvent MouseEnter
            hTimer.Enabled = True
        'the mouse is down and the mouse hasn't been drawn down
        ElseIf (Not m_ButtonState = 4) And (MouseDown = True) Then
            m_ButtonState = DrawButton(4)
        End If
    ElseIf (m_ButtonState = 4) And (MouseDown = True) Then
        'the mouse is down and leaving the control
        m_ButtonState = DrawButton(3)
        MouseOver = False
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    If x >= 0 And x <= UserControl.ScaleWidth And y >= 0 And y <= UserControl.ScaleHeight Then
        If Button = vbLeftButton And MouseDown = True Then
            m_ButtonState = DrawButton(1)
            MouseDown = False
            RaiseEvent Click
            Call GetCursorPos(mPos)
            If Not (WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd) Then
                m_ButtonState = DrawButton(2)
                MouseOver = False
            Else
                m_ButtonState = DrawButton(3)
                MouseOver = True
                hTimer.Enabled = True
            End If
        End If
    Else
        'the mouse was down, but was unclicked outside of the button.
        If Button = vbLeftButton And MouseDown = True Then
            m_ButtonState = DrawButton(2)
            MouseDown = False
        End If
    End If
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_AccessKey = PropBag.ReadProperty("AccessKey", vbNullString)
    m_BackColour = PropBag.ReadProperty("BackColour", m_def_BackColour)
    m_Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ForeColour = PropBag.ReadProperty("ForeColour", m_def_ForeColour)
    m_Picture = PropBag.ReadProperty("Picture", vbNullString)
    m_PictureAlignment = PropBag.ReadProperty("PictureAlignment", m_def_PictureAlignment)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.Size)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
        
    'the default and previous values have been loaded so now draw button
    fInit = True
    UserControl.AccessKeys = m_AccessKey
    DrawButton m_ButtonState
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AccessKey", m_AccessKey, vbNullString)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", m_Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ForeColour", m_ForeColour, m_def_ForeColour)
    Call PropBag.WriteProperty("BackColour", m_BackColour, m_def_BackColour)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("Picture", m_Picture, vbNullString)
    Call PropBag.WriteProperty("PictureAlignment", m_PictureAlignment, m_def_PictureAlignment)
End Sub

Private Sub UserControl_Resize()
    'min size
    If UserControl.ScaleWidth < 10 Then
        UserControl.Width = ScaleX(15, vbPixels, vbTwips)
    ElseIf UserControl.ScaleHeight < 10 Then
        UserControl.Height = ScaleY(10, vbPixels, vbTwips)
    End If
    
    'on resize the gradients must be recalculated
    ReDim GradBack(0)
    ReDim GradFoc(0)
    ReDim GradHot(0)
    DrawButton m_ButtonState
End Sub

Private Sub UserControl_Terminate()
    Set hTimer = Nothing
End Sub

Private Function GetRGB(ByVal LongValue As Long) As RGBType
    LongValue = Abs(LongValue)
    GetRGB.Red = LongValue And 255
    GetRGB.Green = (LongValue \ 256) And 255
    GetRGB.Blue = (LongValue \ 65536) And 255
End Function

'this function draws a line quicker than the Visual Basic line routine.
Private Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, crColour As Long)
    Dim hPen As Long
    Dim hOldPen As Long
    
    hPen = CreatePen(PS_SOLID, 1, crColour)
    hOldPen = SelectObject(UserControl.hdc, hPen)
    
    'set the forecolour before drawing the line. This method is quicker
    'than using a pen.
    'UserControl.ForeColor = crColour
    'move the start of the line to this position
    MoveToEx UserControl.hdc, X1, Y1, ByVal 0&
    'draw the line to this position
    LineTo UserControl.hdc, X2, Y2
    
    SelectObject UserControl.hdc, hOldPen
    DeleteObject hPen
End Sub

Private Sub DrawTextTohWnd(htext As String, lentext As Long)
    Dim vh      As Long 'vertical height of text (wrapped)
    Dim hRect   As RECT 'text boundaries
    
    'set the rectangular area that the text can drawn onto
    SetRect hRect, 4, 0, ScaleWidth - 4, ScaleHeight
    'do a test draw (not actuallt drawn on screen) and find the height
    'the text occupies with text wrapping.
    vh = DrawText(UserControl.hdc, htext, lentext, hRect, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    'set the rectangular area such that the text is drawn
    'horizontally and vertically centered on the form
    SetRect hRect, 4, (ScaleHeight - vh) / 2, ScaleWidth - 4, ScaleHeight
    'now draw the text
    DrawText UserControl.hdc, htext, lentext, hRect, DT_WORDBREAK Or DT_CENTER
End Sub

'writes all the RGB values of the gradient to an array
'which is drawn in the DrawGrad Sub
'Private Sub CreateGrad(gradient array, displacement, start colour, end colour
Private Sub CreateGrad(Grad() As RGBType, disp As Long, scol As Long, ecol As Long)
    Dim rint    As Single
    Dim gint    As Single
    Dim bint    As Single
    Dim i       As Long
    Dim c       As Long
    Dim col1    As RGBType
    Dim col2    As RGBType
    
    col1 = GetRGB(scol)
    col2 = GetRGB(ecol)
    
    rint = (col1.Red - col2.Red) / (disp - 1)
    gint = (col1.Green - col2.Green) / (disp - 1)
    bint = (col1.Blue - col2.Blue) / (disp - 1)
       
    For i = disp To 1 Step -1
        Grad(i).Red = col1.Red - (c * rint)
        Grad(i).Green = col1.Green - (c * gint)
        Grad(i).Blue = col1.Blue - (c * bint)
        c = c + 1
    Next
End Sub

'Used for drawing the side border gradients using DIBits.
Private Sub DIBDrawVertGrad(hGrad() As RGBType, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional bMirror As Boolean)
    Dim x       As Long
    Dim y       As Long
    Dim bHeight As Long
    Dim bWidth  As Long
    
    'width and height of gradient border
    bWidth = Abs(X2 - X1)
    bHeight = Abs(Y2 - Y1)
    
    'Red=2, Green=1, Blue=0
    ReDim hBits.Bits(3, bWidth - 1, bHeight - 1)
    
    With hBits.Header
        .biSize = Len(hBits.Header)
        .biWidth = bWidth
        .biHeight = bHeight
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = 3 * bHeight * bWidth
    End With
    
    For y = 0 To bHeight - 1
        For x = 0 To bWidth - 1
            'create gradient border in array
            hBits.Bits(2, x, y) = hGrad(y + 1).Red
            hBits.Bits(1, x, y) = hGrad(y + 1).Green
            hBits.Bits(0, x, y) = hGrad(y + 1).Blue
        Next
    Next
    
    'draw gradient to screen
    SetDIBitsToDevice UserControl.hdc, X1, Y1, bWidth, bHeight, 0, 0, 0, bHeight, hBits.Bits(0, 0, 0), hBits, 0&
    
    'create mirror DIB
    If bMirror = True Then
        SetDIBitsToDevice UserControl.hdc, UserControl.ScaleWidth - X1 - bWidth, Y1, bWidth, bHeight, 0, 0, 0, bHeight, hBits.Bits(0, 0, 0), hBits, 0&
    End If
    
End Sub

Private Sub FillCol(ucol As Long, fillmode As Byte)
    Dim l       As Byte
    Dim i       As Integer
    Dim hBrush  As Long
    Dim hRect   As RECT
    Dim lPos    As POINTAPI
    
    'solid fill, no gradient
    If fillmode = 1 Then
        Call SetRect(hRect, 1, 1, ScaleWidth - 1, ScaleHeight - 1)
        hBrush = CreateSolidBrush(ucol)
        Call FillRect(UserControl.hdc, hRect, hBrush)
        DeleteObject hBrush
    'gradient fill
    ElseIf (fillmode = 2) Or (fillmode = 3) Then
        'for drawidle, draws extra lines on button
        If fillmode = 3 Then
            DrawLine 3, 1, UserControl.ScaleWidth - 3, 1, TopFirstI
            DrawLine 2, 2, UserControl.ScaleWidth - 2, 2, TopSecondI
            l = 1
        Else 'fillmode = 2
            l = 3
        End If
        
        'draws line at bottom of gradient
        DrawLine CLng(l), UserControl.ScaleHeight - 4, UserControl.ScaleWidth - l, UserControl.ScaleHeight - 4, EndBGrad
            
        If LBound(GradBack) = 0 Then
            ReDim GradBack(1 To UserControl.ScaleHeight - 7)
            'the bottom background gradient and top background gradient are
            'reversed in the function call below. I am not sure why, but it works.
            CreateGrad GradBack, UserControl.ScaleHeight - 7, BotBGrad, TopBGrad
        End If
        
        For i = 1 To UserControl.ScaleHeight - 7
            DrawLine CLng(l), i + 2, UserControl.ScaleWidth - l, i + 2, RGB(GradBack(i).Red, GradBack(i).Green, GradBack(i).Blue)
        Next
    End If
End Sub

Private Sub DrawFocus()
    Dim hfillmode   As Byte
    
    With UserControl
        Call SetColourScheme(1, m_BorderStyle)
        
        If m_BackStyle = 2 Then
            hfillmode = 2
        Else
            hfillmode = 1
        End If
        
        If LowCol = False Then
            If m_BackColour = 1 Then
                FillCol m_def_BackColour_light, hfillmode
            ElseIf m_BackColour = 2 Then
                FillCol m_def_BackColour_med, hfillmode
            Else
                FillCol m_def_BackColour_dark, hfillmode
            End If
        Else
            'don't use gradient background if screen 256 or less colours.
            FillCol m_def_BackColourLow, 1
        End If
        
        'borders
        DrawLine 2, 0, .ScaleWidth - 2, 0, OutBorder
        DrawLine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, OutBorder
        DrawLine 0, 2, 0, .ScaleHeight - 2, OutBorder
        DrawLine .ScaleWidth - 1, 2, .ScaleWidth - 1, .ScaleHeight - 2, OutBorder
        
        DrawLine 2, 1, .ScaleWidth - 2, 1, TopFirst
        DrawLine 3, 2, .ScaleWidth - 3, 2, TopSecond
        
        DrawLine 3, .ScaleHeight - 3, .ScaleWidth - 3, .ScaleHeight - 3, BotFirst
        DrawLine 2, .ScaleHeight - 2, .ScaleWidth - 2, .ScaleHeight - 2, BotSecond
        
        If LBound(GradFoc) = 0 Then
            ReDim GradFoc(1 To .ScaleHeight - 4)
            'create side borders which have a gradient fade unlike the
            'top and bottom borders.
            Call CreateGrad(GradFoc(), .ScaleHeight - 4, TopSecond, BotFirst)
        End If
                    
        'now time to draw the gradient
        Call DIBDrawVertGrad(GradFoc(), 1, 2, 3, .ScaleHeight - 2, True)
        
        'set pixels top left corner
        SetPixelV .hdc, 0, 1, OutPixelF
        SetPixelV .hdc, 1, 0, OutPixelF
        SetPixelV .hdc, 1, 1, InPixelF
        
        'set pixels bottom left corner
        SetPixelV .hdc, 0, .ScaleHeight - 2, OutPixelF
        SetPixelV .hdc, 1, .ScaleHeight - 1, OutPixelF
        SetPixelV .hdc, 1, .ScaleHeight - 2, InPixelF
        
        'set pixels top right corner
        SetPixelV .hdc, .ScaleWidth - 2, 0, OutPixelF
        SetPixelV .hdc, .ScaleWidth - 1, 1, OutPixelF
        SetPixelV .hdc, .ScaleWidth - 2, 1, InPixelF
        
        'set pixels bottom right corner
        SetPixelV .hdc, .ScaleWidth - 1, .ScaleHeight - 2, OutPixelF
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 1, OutPixelF
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, InPixelF
    End With
End Sub

Private Sub DrawMouseIn()
    Dim hfillmode   As Byte

    With UserControl
        Call SetColourScheme(2, m_BorderStyle)
        
        If m_BackStyle = 2 Then
            hfillmode = 2
        Else
            hfillmode = 1
        End If
        
        If LowCol = False Then
            If m_BackColour = 1 Then
                FillCol m_def_BackColour_light, hfillmode
            ElseIf m_BackColour = 2 Then
                FillCol m_def_BackColour_med, hfillmode
            Else
                FillCol m_def_BackColour_dark, hfillmode
            End If
        Else
            'don't use gradient background if screen 256 or less colours.
            FillCol m_def_BackColourLow, 1
        End If
        
        'borders
        DrawLine 2, 0, .ScaleWidth - 2, 0, OutBorder
        DrawLine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, OutBorder
        DrawLine 0, 2, 0, .ScaleHeight - 2, OutBorder
        DrawLine .ScaleWidth - 1, 2, .ScaleWidth - 1, .ScaleHeight - 2, OutBorder
        
        DrawLine 2, 1, .ScaleWidth - 2, 1, TopFirst
        DrawLine 3, 2, .ScaleWidth - 3, 2, TopSecond
        
        DrawLine 3, .ScaleHeight - 3, .ScaleWidth - 3, .ScaleHeight - 3, BotFirst
        DrawLine 2, .ScaleHeight - 2, .ScaleWidth - 2, .ScaleHeight - 2, BotSecond
        
        If LBound(GradHot) = 0 Then
            ReDim GradHot(1 To .ScaleHeight - 4)
            'create side borders which have a gradient fade unlike the
            'top and bottom borders.
            Call CreateGrad(GradHot(), .ScaleHeight - 4, TopSecond, BotFirst)
        End If
        
        'draw the gradient
        Call DIBDrawVertGrad(GradHot(), 1, 2, 3, .ScaleHeight - 2, True)
        
        'set pixels top left corner
        SetPixelV .hdc, 0, 1, OutPixelM
        SetPixelV .hdc, 1, 0, OutPixelM
        SetPixelV .hdc, 1, 1, InPixelM
        
        'set pixels bottom left corner
        SetPixelV .hdc, 0, .ScaleHeight - 2, OutPixelM
        SetPixelV .hdc, 1, .ScaleHeight - 1, OutPixelM
        SetPixelV .hdc, 1, .ScaleHeight - 2, InPixelM
        
        'set pixels top right corner
        SetPixelV .hdc, .ScaleWidth - 2, 0, OutPixelM
        SetPixelV .hdc, .ScaleWidth - 1, 1, OutPixelM
        SetPixelV .hdc, .ScaleWidth - 2, 1, InPixelM
        
        'set pixels bottom right corner
        SetPixelV .hdc, .ScaleWidth - 1, .ScaleHeight - 2, OutPixelM
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 1, OutPixelM
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, InPixelM
    End With
End Sub

Private Sub DrawIdle()
    Dim hfillmode   As Byte

    With UserControl
        If m_BackStyle = 2 Then
            hfillmode = 3
        Else
            hfillmode = 1
        End If
        
        If LowCol = False Then
            If m_BackColour = 1 Then
                FillCol m_def_BackColour_light, hfillmode
            ElseIf m_BackColour = 2 Then
                FillCol m_def_BackColour_med, hfillmode
            Else
                FillCol m_def_BackColour_dark, hfillmode
            End If
        Else
            'don't use gradient background if screen 256 or less colours.
            FillCol m_def_BackColourLow, 1
        End If
        
        'borders
        DrawLine 2, 0, .ScaleWidth - 2, 0, OutBorder
        DrawLine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, OutBorder
        DrawLine 0, 2, 0, .ScaleHeight - 2, OutBorder
        DrawLine .ScaleWidth - 1, 2, .ScaleWidth - 1, .ScaleHeight - 2, OutBorder
        
        'bottom fade
        DrawLine 2, .ScaleHeight - 3, .ScaleWidth - 2, .ScaleHeight - 3, BotFirstI
        DrawLine 3, .ScaleHeight - 2, .ScaleWidth - 3, .ScaleHeight - 2, BotSecondI
        
        'set pixels top left corner
        SetPixelV .hdc, 0, 1, OutPixelI
        SetPixelV .hdc, 1, 0, OutPixelI
        SetPixelV .hdc, 1, 1, MidPixelI
        SetPixelV .hdc, 1, 2, InPixelI
        SetPixelV .hdc, 2, 1, InPixelI
        
        'set pixels top right corner
        SetPixelV .hdc, .ScaleWidth - 2, 0, OutPixelI
        SetPixelV .hdc, .ScaleWidth - 1, 1, OutPixelI
        SetPixelV .hdc, .ScaleWidth - 2, 1, MidPixelI
        SetPixelV .hdc, .ScaleWidth - 3, 1, InPixelI
        SetPixelV .hdc, .ScaleWidth - 2, 2, InPixelI
        
        'set pixels bottom left corner
        SetPixelV .hdc, 0, .ScaleHeight - 2, OutPixelI
        SetPixelV .hdc, 1, .ScaleHeight - 1, OutPixelI
        SetPixelV .hdc, 1, .ScaleHeight - 2, MidPixelI
        SetPixelV .hdc, 1, .ScaleHeight - 3, InPixelI
        SetPixelV .hdc, 2, .ScaleHeight - 2, InPixelI
        
        'set pixels bottom right corner
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 1, OutPixelI
        SetPixelV .hdc, .ScaleWidth - 1, .ScaleHeight - 2, OutPixelI
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, MidPixelI
        SetPixelV .hdc, .ScaleWidth - 3, .ScaleHeight - 2, InPixelI
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 3, InPixelI
    End With
End Sub

Private Sub DrawDown()
    With UserControl
        If LowCol = False Then
            FillCol m_def_BackColourDown, 1
        Else
            FillCol m_def_BackColourDownLow, 1
        End If
        
        'borders
        DrawLine 2, 0, .ScaleWidth - 2, 0, OutBorder
        DrawLine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, OutBorder
        DrawLine 0, 2, 0, .ScaleHeight - 2, OutBorder
        DrawLine .ScaleWidth - 1, 2, .ScaleWidth - 1, .ScaleHeight - 2, OutBorder
        
        'top fade
        DrawLine 3, 1, .ScaleWidth - 3, 1, TopFirstMD
        DrawLine 2, 2, .ScaleWidth - 2, 2, TopSecondMD
        DrawLine 3, 3, .ScaleWidth - 1, 3, TopThirdMD
        DrawLine 3, 4, .ScaleWidth - 1, 4, TopFourthMD
        
        'bottom fade
        DrawLine 2, .ScaleHeight - 3, .ScaleWidth - 2, .ScaleHeight - 3, BotFirstMD
        DrawLine 3, .ScaleHeight - 2, .ScaleWidth - 3, .ScaleHeight - 2, BotSecondMD
        
        'side colours
        DrawLine 1, 3, 1, .ScaleHeight - 3, LeftFirstMD
        DrawLine 2, 3, 2, .ScaleHeight - 3, LeftSecondMD
        
        'set pixels top left corner
        SetPixelV .hdc, 0, 1, OutPixelMD
        SetPixelV .hdc, 1, 0, OutPixelMD
        SetPixelV .hdc, 1, 1, MidPixelMD
        SetPixelV .hdc, 1, 2, InPixelMD
        SetPixelV .hdc, 2, 1, InPixelMD
        
        'set pixels top right corner
        SetPixelV .hdc, .ScaleWidth - 2, 0, OutPixelMD
        SetPixelV .hdc, .ScaleWidth - 1, 1, OutPixelMD
        SetPixelV .hdc, .ScaleWidth - 2, 1, MidPixelMD
        SetPixelV .hdc, .ScaleWidth - 3, 1, InPixelMD
        SetPixelV .hdc, .ScaleWidth - 2, 2, InPixelMD
        
        'set pixels bottom left corner
        SetPixelV .hdc, 0, .ScaleHeight - 2, OutPixelMD
        SetPixelV .hdc, 1, .ScaleHeight - 1, OutPixelMD
        SetPixelV .hdc, 1, .ScaleHeight - 2, MidPixelMD
        SetPixelV .hdc, 1, .ScaleHeight - 3, InPixelMD
        SetPixelV .hdc, 2, .ScaleHeight - 2, InPixelMD
        
        'set pixels bottom right corner
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 1, OutPixelMD
        SetPixelV .hdc, .ScaleWidth - 1, .ScaleHeight - 2, OutPixelMD
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, MidPixelMD
        SetPixelV .hdc, .ScaleWidth - 3, .ScaleHeight - 2, InPixelMD
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 3, InPixelMD
    End With
End Sub

Private Sub DrawDisabled()
    With UserControl
        If LowCol = False Then
            FillCol m_def_DisabledBackColour, 1
        Else
            FillCol m_def_DisableBackColourLow, 1
        End If
        
        'borders
        DrawLine 2, 0, .ScaleWidth - 2, 0, OutBorderD
        DrawLine 2, .ScaleHeight - 1, .ScaleWidth - 2, .ScaleHeight - 1, OutBorderD
        DrawLine 0, 2, 0, .ScaleHeight - 2, OutBorderD
        DrawLine .ScaleWidth - 1, 2, .ScaleWidth - 1, .ScaleHeight - 2, OutBorderD
        
        'set pixels top left corner
        SetPixelV .hdc, 0, 1, OutPixelD
        SetPixelV .hdc, 1, 0, OutPixelD
        SetPixelV .hdc, 1, 1, MidPixelD
        SetPixelV .hdc, 1, 2, InPixelD
        SetPixelV .hdc, 2, 1, InPixelD
        
        'set pixels top right corner
        SetPixelV .hdc, .ScaleWidth - 2, 0, OutPixelD
        SetPixelV .hdc, .ScaleWidth - 1, 1, OutPixelD
        SetPixelV .hdc, .ScaleWidth - 2, 1, MidPixelD
        SetPixelV .hdc, .ScaleWidth - 3, 1, InPixelD
        SetPixelV .hdc, .ScaleWidth - 2, 2, InPixelD
        
        'set pixels bottom left corner
        SetPixelV .hdc, 0, .ScaleHeight - 2, OutPixelD
        SetPixelV .hdc, 1, .ScaleHeight - 1, OutPixelD
        SetPixelV .hdc, 1, .ScaleHeight - 2, MidPixelD
        SetPixelV .hdc, 1, .ScaleHeight - 3, InPixelD
        SetPixelV .hdc, 2, .ScaleHeight - 2, InPixelD
        
        'set pixels bottom right corner
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 1, OutPixelD
        SetPixelV .hdc, .ScaleWidth - 1, .ScaleHeight - 2, OutPixelD
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 2, MidPixelD
        SetPixelV .hdc, .ScaleWidth - 3, .ScaleHeight - 2, InPixelD
        SetPixelV .hdc, .ScaleWidth - 2, .ScaleHeight - 3, InPixelD
    End With
End Sub

Private Sub SetColourScheme(state As Byte, bstyle As Byte)
    If state = 1 Then
        'xp border scheme
        Select Case bstyle
            Case 1
                TopFirst = TopFirstF1
                TopSecond = TopSecondF1
                BotFirst = BotFirstF1
                BotSecond = BotSecondF1
            Case 2
                TopFirst = TopFirstF2
                TopSecond = TopSecondF2
                BotFirst = BotFirstF2
                BotSecond = BotSecondF2
            Case 3
                TopFirst = TopFirstF3
                TopSecond = TopSecondF3
                BotFirst = BotFirstF3
                BotSecond = BotSecondF3
            Case 4
                TopFirst = TopFirstF4
                TopSecond = TopSecondF4
                BotFirst = BotFirstF4
                BotSecond = BotSecondF4
            Case 5
                TopFirst = TopFirstF5
                TopSecond = TopSecondF5
                BotFirst = BotFirstF5
                BotSecond = BotSecondF5
            Case 6
                TopFirst = TopFirstF6
                TopSecond = TopSecondF6
                BotFirst = BotFirstF6
                BotSecond = BotSecondF6
            Case 7
                TopFirst = TopFirstF7
                TopSecond = TopSecondF7
                BotFirst = BotFirstF7
                BotSecond = BotSecondF7
            Case 8
                TopFirst = TopFirstF8
                TopSecond = TopSecondF8
                BotFirst = BotFirstF8
                BotSecond = BotSecondF8
        End Select
        
    ElseIf state = 2 Then
        
        'xp border scheme
        Select Case bstyle
            Case 1
                TopFirst = TopFirstM1
                TopSecond = TopSecondM1
                BotFirst = BotFirstM1
                BotSecond = BotSecondM1
            Case 2
                TopFirst = TopFirstM2
                TopSecond = TopSecondM2
                BotFirst = BotFirstM2
                BotSecond = BotSecondM2
            Case 3
                TopFirst = TopFirstM3
                TopSecond = TopSecondM3
                BotFirst = BotFirstM3
                BotSecond = BotSecondM3
            Case 4
                TopFirst = TopFirstM4
                TopSecond = TopSecondM4
                BotFirst = BotFirstM4
                BotSecond = BotSecondM4
            Case 5
                TopFirst = TopFirstM5
                TopSecond = TopSecondM5
                BotFirst = BotFirstM5
                BotSecond = BotSecondM5
            Case 6
                TopFirst = TopFirstM6
                TopSecond = TopSecondM6
                BotFirst = BotFirstM6
                BotSecond = BotSecondM6
            Case 7
                TopFirst = TopFirstM7
                TopSecond = TopSecondM7
                BotFirst = BotFirstM7
                BotSecond = BotSecondM7
            Case 8
                TopFirst = TopFirstM8
                TopSecond = TopSecondM8
                BotFirst = BotFirstM8
                BotSecond = BotSecondM8
        End Select
    End If
End Sub

Private Sub SetAccessKey()
    On Error GoTo NoAccessKeys
    Dim i   As Integer
    
    m_AccessKey = vbNullString
    i = Len(m_Caption)

    Do
        i = InStrRev(m_Caption, "&", i)
        
        If i = 1 Then
            GoTo SkipCheck
        ElseIf i = Len(m_Caption) Then
            GoTo NoAccessKeys
        End If
        
        If Mid$(m_Caption, i - 1, 1) = "&" Then
            i = i - 2
        Else
SkipCheck:
            m_AccessKey = Mid$(m_Caption, i + 1, 1)
        End If
        
    Loop Until (m_AccessKey <> vbNullString) Or (i = 0)
    
    UserControl.AccessKeys = m_AccessKey

Exit Sub
NoAccessKeys:
m_AccessKey = vbNullString
UserControl.AccessKeys = vbNullString
End Sub

Private Function DrawButton(lstate As Byte, Optional cornerpixels As Boolean) As Byte
    If fInit = False Then Exit Function

    'if the control is disabled then always draw disabled.
    If UserControl.Enabled = False Then
        Call DrawDisabled
    Else
        Select Case lstate
            Case 1
                Call DrawIdle
            Case 2
                Call DrawFocus
            Case 3
                Call DrawMouseIn
            Case 4
                Call DrawDown
        End Select
    End If
        
    If Not (m_Caption = "") Then
        If UserControl.Enabled = True Then
            UserControl.ForeColor = m_ForeColour
        Else
            UserControl.ForeColor = ForeColourD 'disabled forecolour
        End If
        
        'draws the caption to the usercontrol
        DrawTextTohWnd m_Caption, Len(m_Caption)
    End If
    
    'redraw the window
    RedrawWindow UserControl.hwnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE

    DrawButton = lstate
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If UserControl.Enabled() = New_Enabled Then Exit Property
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    'If the usercontrol was disabled and now is enabled
    If UserControl.Enabled() = True Then
        Call GetCursorPos(mPos)
        'cursor is inside usercontrol
        If WindowFromPoint(mPos.x, mPos.y) = UserControl.hwnd Then
            'draw mousein and enable 'mouse checking timer'.
            m_ButtonState = DrawButton(3)
            hTimer.Enabled = True
        Else 'curosr is outside of usercontrol
            'draw normal
            m_ButtonState = DrawButton(1)
        End If
    Else 'the usercontrol was enabled and is now disabled
        m_ButtonState = DrawButton(1)
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    DrawButton m_ButtonState
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    If m_Caption = New_Caption Then Exit Property
    m_Caption = New_Caption
    PropertyChanged "Caption"
    SetAccessKey
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get BorderStyle() As BorderStyleEnum
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
    If New_BorderStyle = m_BorderStyle Then Exit Property
    If New_BorderStyle > 8 Or New_BorderStyle < 1 Then Exit Property
        
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    
    'new gradient must be drawn
    ReDim GradFoc(0)
    ReDim GradHot(0)
    
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    If UserControl.FontBold() = New_FontBold Then Exit Property
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    If UserControl.FontItalic() = New_FontItalic Then Exit Property
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    If UserControl.FontName() = New_FontName Then Exit Property
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    If UserControl.FontSize() = New_FontSize Then Exit Property
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    If UserControl.FontStrikethru() = New_FontStrikethru Then Exit Property
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    If UserControl.FontUnderline() = New_FontUnderline Then Exit Property
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,00
Public Property Get ForeColour() As OLE_COLOR
    ForeColour = m_ForeColour
End Property

Public Property Let ForeColour(ByVal New_ForeColour As OLE_COLOR)
    If m_ForeColour = New_ForeColour Then Exit Property
    m_ForeColour = New_ForeColour
    PropertyChanged "ForeColour"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,2
Public Property Get BackColour() As BackColourEnum
    BackColour = m_BackColour
End Property

Public Property Let BackColour(ByVal New_BackColour As BackColourEnum)
    If m_BackColour = New_BackColour Then Exit Property
    If New_BackColour > 3 Or New_BackColour < 1 Then Exit Property
    
    m_BackColour = New_BackColour
    PropertyChanged "BackColour"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get BackStyle() As BackStyleEnum
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As BackStyleEnum)
    If m_BackStyle = New_BackStyle Then Exit Property
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get PictureAlignment() As PictureAlignmentEnum
    PictureAlignment = m_PictureAlignment
End Property

Public Property Let PictureAlignment(ByVal New_PictureAlignment As PictureAlignmentEnum)
    If New_PictureAlignment = m_PictureAlignment Then Exit Property
    m_PictureAlignment = New_PictureAlignment
    PropertyChanged "PictureAlignment"
    DrawButton m_ButtonState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Picture() As String
    Picture = m_Picture
End Property

Public Property Let Picture(ByVal New_Picture As String)
    If New_Picture = m_Picture Then Exit Property
    m_Picture = New_Picture
    PropertyChanged "Picture"
End Property
