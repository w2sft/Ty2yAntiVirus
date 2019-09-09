VERSION 5.00
Begin VB.UserControl Command 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1815
      Top             =   1185
   End
End
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTERABS = &H65                    'CENTERABS= &H65
Const DT_WORDBREAK = &H10
Const DT_CENTER = &H1
Const DT_VCENTER = &H4
Const DT_EXPANDTABS = &H40
Const DT_EXTERNALLEADING = &H200
Const DT_CALCRECT = &H400

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


'Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long


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

Public Enum ButtonTypes
    [Windows XP] = 1        'the new brand XP button totally owner-drawn
    [Mac] = 2               'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
    [Mac OS] = 3
    [Longhorn] = 4
    [Office Xp] = 5
End Enum

Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
    [Force Standard] = 3
End Enum
Public Enum XpTypes
    [银色风格] = 1
    [翠色风格] = 2
    [蓝色风格] = 3
End Enum
'variables
'属性变量:
'Dim m_ButtonType As ButtonTypes
'Dim m_ColorScheme As ColorTypes

Dim m_rectcolor As OLE_COLOR
Dim m_hWnd As Long
Dim showFocusR As Boolean
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes
Private MyXpType As XpTypes

Private He As Long  'wwbutton的高度
Private Wi As Long  'wwbutton的宽度
Dim allcount As Integer                       'caption字节总数
Private BackC As Long 'back color
Private ForeC As Long 'fore color

Private m_Caption As String     'current caption  变量
Private TextFont As StdFont 'current font

Private Rc As RECT, rc2 As RECT, rc3 As RECT
Private rgnNorm As Long                           '正常区域句柄

Private LastButton As Byte, LastKeyDown As Byte         '上一次按钮状态和上一次键盘按下状态
Private isEnabled As Boolean
Private hasFocus As Boolean                         '焦点标志
Private disyellowrect As Boolean                         '鼠标移入时显示黄色圆角矩形标志
Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long

Dim m_Percent As Integer                '按钮上边沿到按钮上颜色最深的位置之距离占按钮高度的百分比
Private m_MidColor As Long                             '以下六个变量为Mac OS和Longhorn中间色MidColor和边色EndColor
Private m_EndColor As Long
Private m_MouseMoveMidColor As Long
Private m_MouseMoveEndColor As Long
Private m_MouseDownMidColor As Long
Private m_MouseDownEndColor As Long

Private m_OfficeXpFillColor As Long                  '正常状态内部填充色
Private m_OfficeXpMousemoveFillColor As Long         '鼠标移入时内部填充色
Private m_OfficeXpFrameColor As Long                 '正常状态外框色
Private m_OfficeXpMousemoveFrameColor As Long        '鼠标移入时外框色
Private lastStat As Byte, TE As String '        保存状态，消除不必要的重画
'缺省属性值:


Const m_def_rectcolor = &H58C3FA

Const m_def_Enabled = True
Const m_def_hWnd = 0
Const m_def_ButtonType = [Windows XP]
Const m_def_XpType = [银色风格]
Const m_def_Caption = "wwbutton"
Const m_def_ColorScheme = [Use Windows]
Const m_def_ShowFocusRect = True

Const m_def_Percent = 16
Const m_def_MidColor = &H73C874               'RGB(116, 200, 115)
Const m_def_EndColor = &HD7F6D7                 'RGB(215, 246, 215)
Const m_def_MouseMoveMidColor = &H33A335                   'RGB(53, 163, 51)
Const m_def_MouseMoveEndColor = &HF4FCF5         'RGB(245, 252, 244)
Const m_def_MouseDownMidColor = &H85D084         'RGB(132, 208, 133)
Const m_def_MouseDownEndColor = &H0
Const m_def_OfficeXpFillColor = &HA0DEDE
Const m_def_OfficeXpFrameColor = &H42BCBC
Const m_def_OfficeXpMousemoveFillColor = &HBADCDC
Const m_def_OfficeXpMousemoveFrameColor = &H1807F

'事件声明:

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseOut()


'********************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = BackC
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    BackC = New_BackColor
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "BackColor"
End Property

'***********************************************************************************
'********************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = ForeC
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    ForeC = New_ForeColor
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "ForeColor"
End Property

'Mac OS Property
Public Property Get Percent() As Integer              'Percent设置按钮从上到下中间色的位置，以百分比计
    Percent = m_Percent
End Property

Public Property Let Percent(ByVal New_Percent As Integer)
    m_Percent = New_Percent
    If m_Percent < 0 Or m_Percent > 100 Then
       Err.Raise 380          ' MsgBox "无效属性"
    Else
       Call Redraw(lastStat, True)
       PropertyChanged "Percent"
    End If
    'Call SetColors
End Property
'***************************************************************************
'以下为Mac OS和Longhorn风格属性
Public Property Get MidColor() As OLE_COLOR
    MidColor = m_MidColor
End Property

Public Property Let MidColor(ByVal New_MidColor As OLE_COLOR)            'MidColor
    m_MidColor = New_MidColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MidColor"
End Property
'*************************************************************************
Public Property Get EndColor() As OLE_COLOR
    EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal New_EndColor As OLE_COLOR)            'EndColor
    m_EndColor = New_EndColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "EndColor"
End Property
'***********************************************************************
Public Property Get MouseMoveMidColor() As OLE_COLOR
    MouseMoveMidColor = m_MouseMoveMidColor
End Property

Public Property Let MouseMoveMidColor(ByVal New_MouseMoveMidColor As OLE_COLOR)            'MouseMoveMidColor
    m_MouseMoveMidColor = New_MouseMoveMidColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MouseMoveMidColor"
End Property
'*************************************************************************
Public Property Get MouseMoveEndColor() As OLE_COLOR
    MouseMoveEndColor = m_MouseMoveEndColor
End Property

Public Property Let MouseMoveEndColor(ByVal New_MouseMoveEndColor As OLE_COLOR)            'MouseMoveEndColor
    m_MouseMoveEndColor = New_MouseMoveEndColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MouseMoveEndColor"
End Property
'****************************************************************************
Public Property Get MouseDownMidColor() As OLE_COLOR
    MouseDownMidColor = m_MouseDownMidColor
End Property

Public Property Let MouseDownMidColor(ByVal New_MouseDownMidColor As OLE_COLOR)            'MouseDownMidColor
    m_MouseDownMidColor = New_MouseDownMidColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MouseDownMidColor"
End Property
'*****************************************************************************
Public Property Get MouseDownEndColor() As OLE_COLOR
    MouseDownEndColor = m_MouseDownEndColor
End Property

Public Property Let MouseDownEndColor(ByVal New_MouseDownEndColor As OLE_COLOR)            'MouseDownEndColor
    m_MouseDownEndColor = New_MouseDownEndColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MouseDownEndColor"
End Property
Public Property Get OfficeXpFillColor() As OLE_COLOR
    OfficeXpFillColor = m_OfficeXpFillColor
End Property

Public Property Let OfficeXpFillColor(ByVal New_OfficeXpFillColor As OLE_COLOR)            '
    m_OfficeXpFillColor = New_OfficeXpFillColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "OfficeXpFillColor"
End Property

Public Property Get OfficeXpFrameColor() As OLE_COLOR
    OfficeXpFrameColor = m_OfficeXpFrameColor
End Property

Public Property Let OfficeXpFrameColor(ByVal New_OfficeXpFrameColor As OLE_COLOR)
    m_OfficeXpFrameColor = New_OfficeXpFrameColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "OfficeXpFrameColor"
End Property
Public Property Get OfficeXpMousemoveFillColor() As OLE_COLOR
    OfficeXpMousemoveFillColor = m_OfficeXpMousemoveFillColor
End Property

Public Property Let OfficeXpMousemoveFillColor(ByVal New_OfficeXpMousemoveFillColor As OLE_COLOR)
    m_OfficeXpMousemoveFillColor = New_OfficeXpMousemoveFillColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "OfficeXpMousemoveFillColor"
End Property
Public Property Get OfficeXpMousemoveFrameColor() As OLE_COLOR
    OfficeXpMousemoveFrameColor = m_OfficeXpMousemoveFrameColor
End Property

Public Property Let OfficeXpMousemoveFrameColor(ByVal New_OfficeXpMousemoveFrameColor As OLE_COLOR)
    m_OfficeXpMousemoveFrameColor = New_OfficeXpMousemoveFrameColor
    'Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "OfficeXpMousemoveFrameColor"
End Property
'**********************************************************************************
'**************************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    isEnabled = New_Enabled
    'Call Redraw(lastStat, True)                                           '####(0, True)
    UserControl.Enabled = isEnabled
    UserControl_Resize
    PropertyChanged "Enabled"
End Property
'********************************************************************************

'********************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=6,0,0,0
Public Property Get Font() As StdFont
    Set Font = TextFont
End Property

Public Property Set Font(ByVal New_Font As StdFont)
With TextFont
      .Bold = New_Font.Bold
      .Italic = New_Font.Italic
      .Name = New_Font.Name
      .Size = New_Font.Size
End With
    Set TextFont = New_Font
    Set UserControl.Font = TextFont
    Call Redraw(lastStat, True)                                         '####(0, True)
    PropertyChanged "Font"
End Property

'*********************************************************************************
Private Sub UserControl_Initialize()
LastButton = 1
rc2.Left = 2: rc2.Top = 2
Call SetColors
'Set TextFont = New StdFont
'Set UserControl.Font = TextFont
End Sub
'**********************************************************************************
'************************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
'
Public Property Let hWnd(ByVal New_hWnd As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property
''***********************************************************************************

''**************************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=21,0,0,0
Public Property Get ButtonType() As ButtonTypes
    ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal New_ButtonType As ButtonTypes)
    MyButtonType = New_ButtonType
    Call SetColors
    If MyButtonType = 4 Then
                m_Percent = 53
                m_MidColor = RGB(124, 171, 255)
                m_EndColor = RGB(191, 211, 255)
                m_MouseMoveMidColor = RGB(157, 203, 255)
                m_MouseMoveEndColor = RGB(209, 231, 255)
                m_MouseDownMidColor = RGB(100, 153, 255)
                m_MouseDownEndColor = RGB(191, 211, 255)
    End If
    Call UserControl_Resize
    'Call Redraw(lastStat, True)                                         '####(0, True)
    PropertyChanged "ButtonType"
End Property

'**************************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    Call SetAccessKeys
    Call Redraw(lastStat, True)                                         '####(0, True)
    PropertyChanged "Caption"
End Property
'******************************************************************************************
'

''*********************************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=22,0,0,0
Public Property Get ColorScheme() As ColorTypes
    ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal New_ColorScheme As ColorTypes)
    MyColorType = New_ColorScheme
    Call SetColors
    Call Redraw(lastStat, True)                                        '####(0, True)
    PropertyChanged "ColorScheme"
End Property


Public Property Get XpType() As XpTypes
    XpType = MyXpType
End Property

Public Property Let XpType(ByVal New_XpType As XpTypes)
    MyXpType = New_XpType
    'Call SetColors
    Call Redraw(lastStat, True)                                                '####Call Redraw(0, True)
    PropertyChanged "XpType"
End Property
'*********************************************************************************************


'注意！不要删除或修改下列被注释的行！
'MemberInfo=10,0,0,0
Public Property Get rectcolor() As OLE_COLOR
    rectcolor = m_rectcolor
End Property

Public Property Let rectcolor(ByVal New_rectcolor As OLE_COLOR)
    m_rectcolor = New_rectcolor
    PropertyChanged "rectcolor"
End Property
'***********************************************************************************************
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = showFocusR
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    showFocusR = New_ShowFocusRect
    Call Redraw(lastStat, True)
    PropertyChanged "ShowFocusRect"
End Property
'***********************************************************************************************
'为用户控件初始化属性
Private Sub UserControl_InitProperties()

    BackC = GetSysColor(COLOR_BTNFACE)
    ForeC = GetSysColor(COLOR_BTNTEXT)
    isEnabled = m_def_Enabled
   ' Set TextFont = Ambient.Font
    Set TextFont = UserControl.Font
    m_hWnd = m_def_hWnd
    MyButtonType = m_def_ButtonType
    m_Caption = Extender.Name                          ' m_def_Caption
    MyColorType = m_def_ColorScheme
    MyXpType = [银色风格]
    showFocusR = m_def_ShowFocusRect
    m_rectcolor = m_def_rectcolor
    lastStat = 0
    
    m_Percent = m_def_Percent
    m_MidColor = m_def_MidColor
    m_EndColor = m_def_EndColor
    m_MouseMoveMidColor = m_def_MouseMoveMidColor
    m_MouseMoveEndColor = m_def_MouseMoveEndColor
    m_MouseDownMidColor = m_def_MouseDownMidColor
    m_MouseDownEndColor = m_def_MouseDownEndColor
    m_OfficeXpFillColor = m_def_OfficeXpFillColor
    m_OfficeXpFrameColor = m_def_OfficeXpFrameColor
    m_OfficeXpMousemoveFillColor = m_def_OfficeXpMousemoveFillColor
    m_OfficeXpMousemoveFrameColor = m_def_OfficeXpMousemoveFrameColor
    'm_TwoState = m_def_TwoState
    'm_Value = m_def_Value
End Sub

'***********************************************************************************************
'***********************************************************************************************
'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    BackC = PropBag.ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
    ForeC = PropBag.ReadProperty("ForeColor", GetSysColor(COLOR_BTNTEXT))
    isEnabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set TextFont = PropBag.ReadProperty("Font", Ambient.Font)
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    'm_Value = PropBag.ReadProperty("Value", m_def_Value)
    MyButtonType = PropBag.ReadProperty("ButtonType", m_def_ButtonType)
    MyXpType = PropBag.ReadProperty("XpType", m_def_XpType)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    MyColorType = PropBag.ReadProperty("ColorScheme", m_def_ColorScheme)
    showFocusR = PropBag.ReadProperty("ShowFocusRect", m_def_ShowFocusRect)
    UserControl.Enabled = isEnabled
    Set UserControl.Font = TextFont
    Call SetColors
    Call SetAccessKeys
   
    m_rectcolor = PropBag.ReadProperty("rectcolor", m_def_rectcolor)
    m_Percent = PropBag.ReadProperty("Percent", m_def_Percent)
    m_MidColor = PropBag.ReadProperty("MidColor", m_def_MidColor)
    m_EndColor = PropBag.ReadProperty("EndColor", m_def_EndColor)
    m_MouseMoveEndColor = PropBag.ReadProperty("MouseMoveEndColor", m_def_MouseMoveEndColor)
    m_MouseMoveMidColor = PropBag.ReadProperty("MouseMoveMidColor", m_def_MouseMoveMidColor)
    m_MouseDownEndColor = PropBag.ReadProperty("MouseDownEndColor", m_def_MouseDownEndColor)
    m_MouseDownMidColor = PropBag.ReadProperty("MouseDownMidColor", m_def_MouseDownMidColor)
    m_OfficeXpFillColor = PropBag.ReadProperty("OfficeXpFillColor", m_def_OfficeXpFillColor)
    m_OfficeXpFrameColor = PropBag.ReadProperty("OfficeXpFrameColor", m_def_OfficeXpFrameColor)
    m_OfficeXpMousemoveFillColor = PropBag.ReadProperty("OfficeXpMousemoveFillColor", m_def_OfficeXpMousemoveFillColor)
    m_OfficeXpMousemoveFrameColor = PropBag.ReadProperty("OfficeXpMousemoveFrameColor", m_def_OfficeXpMousemoveFrameColor)

       lastStat = 0

    Call Redraw(lastStat, True)                                   '####(0,true)
    
End Sub
'***********************************************************************************************
'************************************************************************************************
'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", BackC, GetSysColor(COLOR_BTNFACE))
    Call PropBag.WriteProperty("ForeColor", ForeC, GetSysColor(COLOR_BTNTEXT))
    Call PropBag.WriteProperty("Enabled", isEnabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", TextFont, Ambient.Font)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("ButtonType", MyButtonType, m_def_ButtonType)
    Call PropBag.WriteProperty("XpType", MyXpType, m_def_XpType)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("ColorScheme", MyColorType, m_def_ColorScheme)
    Call PropBag.WriteProperty("ShowFocusRect", showFocusR, m_def_ShowFocusRect)

    Call PropBag.WriteProperty("rectcolor", m_rectcolor, m_def_rectcolor)
    Call PropBag.WriteProperty("Percent", m_Percent, m_def_Percent)
    Call PropBag.WriteProperty("MidColor", m_MidColor, m_def_MidColor)
    Call PropBag.WriteProperty("EndColor", m_EndColor, m_def_EndColor)
    Call PropBag.WriteProperty("MouseMoveMidColor", m_MouseMoveMidColor, m_def_MouseMoveMidColor)
    Call PropBag.WriteProperty("MouseMoveEndColor", m_MouseMoveEndColor, m_def_MouseMoveEndColor)
    Call PropBag.WriteProperty("MouseDownMidColor", m_MouseDownMidColor, m_def_MouseDownMidColor)
    Call PropBag.WriteProperty("MouseDownEndColor", m_MouseDownEndColor, m_def_MouseDownEndColor)
    Call PropBag.WriteProperty("OfficeXpFillColor", m_OfficeXpFillColor, m_def_OfficeXpFillColor)
    Call PropBag.WriteProperty("OfficeXpFrameColor", m_OfficeXpFrameColor, m_def_OfficeXpFrameColor)
    Call PropBag.WriteProperty("OfficeXpMousemoveFillColor", m_OfficeXpMousemoveFillColor, m_def_OfficeXpMousemoveFillColor)
    Call PropBag.WriteProperty("OfficeXpMousemoveFrameColor", m_OfficeXpMousemoveFrameColor, m_def_OfficeXpMousemoveFrameColor)

End Sub
'*****************************************************************************************************
'*************************************************************************************************
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
'Debug.Print PropertyName
Call Redraw(lastStat, True)
End Sub


'******************************************************************************************************
Private Sub UserControl_DblClick()
If isEnabled = True Then
   If LastButton = 1 Then
      Call UserControl_MouseDown(1, 1, 1, 1)
   End If
End If
End Sub

Private Sub UserControl_GotFocus()
hasFocus = True
Call Redraw(lastStat, True)
End Sub
'*********************************************************************************************
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If isEnabled = True Then
   RaiseEvent KeyDown(KeyCode, Shift)

   LastKeyDown = KeyCode
   If KeyCode = 32 Then 'spacebar pressed
       Call UserControl_MouseDown(1, 1, 1, 1)
   ElseIf (KeyCode = 39) Or (KeyCode = 40) Then 'right and down arrows
       SendKeys "{Tab}"
   ElseIf (KeyCode = 37) Or (KeyCode = 38) Then 'left and up arrows
       SendKeys "+{Tab}"
   End If
End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If isEnabled = True Then
   RaiseEvent KeyPress(KeyAscii)
End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If isEnabled = True Then
   RaiseEvent KeyUp(KeyCode, Shift)

   If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed
       Call UserControl_MouseUp(1, 1, 1, 1)
       LastButton = 1
       Call UserControl_Click
   End If
End If
End Sub

Private Sub UserControl_LostFocus()
hasFocus = False
Call Redraw(lastStat, True)
End Sub


Private Sub UserControl_Click()
   If isEnabled = True Then
      If (LastButton = 1) Then
          lastStat = 0
          'Call Redraw(lastStat, True)                                '####(0, True) 'be sure that the normal status is drawn

      End If
          Call Redraw(lastStat, True)
          UserControl.Refresh
          RaiseEvent Click
    End If
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If isEnabled = True Then

   'Else
   
      LastButton = Button
      If Button <> 2 Then lastStat = 2                    '####Call Redraw(2, False)
   'End If
   Call Redraw(lastStat, True)
   RaiseEvent MouseDown(Button, Shift, x, y)
End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If isEnabled = True Then

     If Button <> 2 Then lastStat = 0                    '####Call Redraw(0, False)

  Call Redraw(lastStat, True)
  RaiseEvent MouseUp(Button, Shift, x, y)
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button < 2 And isEnabled = True Then
  
    Timer1.Enabled = True
    If x >= 0 And y >= 0 And _
       x <= UserControl.ScaleWidth And y <= UserControl.ScaleHeight Then
       '在按钮内部
        RaiseEvent MouseMove(Button, Shift, x, y)
        If Button = vbLeftButton Then
            Call Redraw(2, False)
            
        Else
             If disyellowrect = True Then
                Exit Sub
             Else
                disyellowrect = True
             End If
             Call Redraw(0, True)
        End If
    End If
  
End If

End Sub

Private Sub UserControl_Resize()
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    Rc.Bottom = He: Rc.Right = Wi
    rc2.Bottom = He: rc2.Right = Wi
    rc3.Left = 2: rc3.Top = 2: rc3.Right = Wi - 2: rc3.Bottom = He - 2
    
    DeleteObject rgnNorm
    Call MakeRegion
    SetWindowRgn UserControl.hWnd, rgnNorm, True
    
    Call Redraw(lastStat, True)                                '####(0, True)
End Sub

Private Sub UserControl_Terminate()
    DeleteObject rgnNorm
End Sub

'按钮重画子程序

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
   Dim I As Long, stepXP1 As Single, XPface As Long
   Dim preFocusValue As Boolean
   Dim lens As Integer
   Dim iunicode As Integer
   Dim ii As Integer, NumLine As Long       'rcText As RECT,
   'Dim pt As POINTAPI
If Force = False Then 'check drawing redundancy
    If (curStat = lastStat) And (TE = m_Caption) Then Exit Sub
End If

If He = 0 Then Exit Sub 'we don't want errors

   lastStat = curStat
   TE = m_Caption
   preFocusValue = hasFocus '保存焦点状态
If hasFocus = True Then hasFocus = ShowFocusRect

With UserControl
     .Cls
     DrawRectangle 0, 0, Wi, He, cFace
     allcount = LenB(StrConv(m_Caption, vbFromUnicode))  '此段为中英文显示判断
     
If isEnabled = True Then
    SetTextColor .hdc, cText 'restore font color
    If curStat = 0 Then
'#@#@#@#@#@# 按钮在正常状态 #@#@#@#@#@#
        
        Select Case MyButtonType

            Case 1 'Windows XP
                 If MyColorType = 2 Or MyColorType = 3 Then
                    stepXP1 = 25 / He
                    XPface = ShiftColor(cFace, &H30, True)
                    For I = 1 To He - 2
                        DrawLine 0, I, Wi, I, ShiftColor(XPface, -stepXP1 * I, True)
                    Next
                  Else
                    Select Case MyXpType
                         Case 1      '银色风格
                              DrawJianBian 198, 197, 215, He - 2, 1, -1 'XP银色效果
                         Case 2               '翠色风格
                              stepXP1 = 25 / He
                              XPface = RGB(242, 237, 218)
                              For I = He - 2 To 1 Step -1
                                  XPface = ShiftColor(XPface, 1, False)
                                  DrawLine 1, I, Wi - 2, I, XPface
                              Next
                         Case 3      '蓝色风格
                               DrawJianBian 236, 235, 230, He - 2, 1, -1   '蓝色风格
                     End Select
                  End If
                  'SetTextColor UserControl.hdc, cText
                  DrawXPFrame &H733C00, &H7B4D10
                If (hasFocus = True) Or ((Ambient.DisplayAsDefault = True) And (showFocusR = True)) Then
                      If disyellowrect = True Then
                           If MyXpType = 2 Then
                              XPface = RGB(228, 144, 80)
                           Else
                              XPface = m_rectcolor
                           End If                                                     '鼠标移入时显示黄色框
                           DrawRectangle 2, 1, Wi - 4, He - 2, XPface, True
                           DrawLine 1, 2, 1, He - 2, XPface
                           DrawLine 3, 2, Wi - 3, 2, XPface
                           DrawLine Wi - 2, 2, Wi - 2, He - 2, XPface
                           DrawLine 3, He - 3, Wi - 3, He - 3, XPface
                      Else
                           Select Case MyXpType
                                  Case 1, 3
                                       DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                                       DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                                       DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                                       DrawLine 2, 2, Wi - 2, 2, &HF7D7BD
                                       DrawLine 2, 3, 2, He - 3, &HF0D1B5
                                       DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                                   Case 2
                                       DrawRectangle 1, 2, Wi - 2, He - 4, &H62B87A, True
                                       DrawLine 2, He - 2, Wi - 2, He - 2, &H62B87A
                                       DrawLine 2, 1, Wi - 2, 1, &H62B87A
                                       DrawLine 2, 2, Wi - 2, 2, &H62B87A
                                       DrawLine 2, 3, 2, He - 3, &H62B87A                           ' FF00&
                                       DrawLine Wi - 3, 3, Wi - 3, He - 3, &H62B87A                 ' &HABF15E
                           End Select
                           DrawFocusR
                      End If
                Else
                     If disyellowrect = True Then
                          If MyXpType = 2 Then                                          '
                             XPface = RGB(228, 144, 80)
                          Else
                             XPface = m_rectcolor '
                          End If                                                        '
                          DrawRectangle 2, 1, Wi - 4, He - 2, XPface, True                '鼠标移入时显示黄色框
                          DrawLine 1, 2, 1, He - 2, XPface                                 '
                          DrawLine 3, 2, Wi - 3, 2, XPface                                 '
                          DrawLine Wi - 2, 2, Wi - 2, He - 2, XPface                      '
                          DrawLine 3, He - 3, Wi - 3, He - 3, XPface                       '
                    
                     End If
                End If
            Case 2 'Mac
                      DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                      DrawLine 2, 0, Wi - 2, 0, cDarkShadow
                      DrawLine 2, He - 1, Wi - 2, He - 1, cDarkShadow
                      DrawLine 0, 2, 0, He - 2, cDarkShadow
                      DrawLine Wi - 1, 2, Wi - 1, He - 2, cDarkShadow
                      mSetPixel 1, 1, cDarkShadow
                      mSetPixel 1, He - 2, cDarkShadow
                      mSetPixel Wi - 2, 1, cDarkShadow
                      mSetPixel Wi - 2, He - 2, cDarkShadow
                      mSetPixel 1, 2, cFace
                      mSetPixel 2, 1, cFace
                      DrawLine 3, 2, Wi - 3, 2, cHighLight
                      DrawLine 2, 2, 2, He - 3, cHighLight
                      mSetPixel 3, 3, cHighLight
                      DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
                      DrawLine 1, He - 3, Wi - 3, He - 3, cFace
                      mSetPixel Wi - 4, He - 4, cFace
                      DrawLine Wi - 2, 3, Wi - 2, He - 2, cShadow
                      DrawLine 3, He - 2, Wi - 2, He - 2, cShadow
                      mSetPixel Wi - 3, He - 3, cShadow
                      mSetPixel 2, He - 2, cFace
                      mSetPixel 2, He - 3, cLight
                      mSetPixel Wi - 2, 2, cFace
                      mSetPixel Wi - 3, 2, cLight
                If (hasFocus = True) Or ((Ambient.DisplayAsDefault = True) And (showFocusR = True)) Then
                    If disyellowrect = True Then
                       DrawRectangle 2, 1, Wi - 4, He - 2, m_rectcolor, True                ''鼠标移入时显示黄色框
                       DrawLine 1, 2, 1, He - 2, m_rectcolor                                 '
                       DrawLine 3, 2, Wi - 3, 2, m_rectcolor                                 '
                       DrawLine Wi - 2, 2, Wi - 2, He - 2, m_rectcolor                       '
                       DrawLine 3, He - 3, Wi - 3, He - 3, m_rectcolor
                    Else
                       DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                       DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                       DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                       DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                       DrawLine 2, 3, 2, He - 3, &HF0D1B5
                       DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                    End If
                Else
                     If disyellowrect = True Then
                        DrawRectangle 2, 1, Wi - 4, He - 2, m_rectcolor, True                ''鼠标移入时显示黄色框
                        DrawLine 1, 2, 1, He - 2, m_rectcolor                                 '
                        DrawLine 3, 2, Wi - 3, 2, m_rectcolor                                 '
                        DrawLine Wi - 2, 2, Wi - 2, He - 2, m_rectcolor                       '
                        DrawLine 3, He - 3, Wi - 3, He - 3, m_rectcolor
                    End If
                End If

             Case 3                                         '[Mac OS] = 8

                  If MyColorType = 1 Then
                      DrawMacOS m_MidColor, RGB(183, 183, 183), m_EndColor, m_Percent
                  Else
                      DrawMacOS m_MidColor, RGB(1, 109, 1), m_EndColor, m_Percent
                  End If
                    If disyellowrect = True Then
                          If MyColorType = 1 Then
                             DrawMacOS m_MouseMoveMidColor, RGB(82, 134, 182), m_MouseMoveEndColor, m_Percent
                          Else
                             DrawMacOS m_MouseMoveMidColor, RGB(1, 109, 1), m_MouseMoveEndColor, m_Percent
                          End If
                         
                     End If
                If (hasFocus = True) Or ((Ambient.DisplayAsDefault = True) And (showFocusR = True)) Then
                     DrawFocusR
                End If
                
             Case 4                                               '[Longhorn] = 4
                    If disyellowrect = True Then
                         DrawLonghorn m_MouseMoveMidColor, RGB(81, 107, 148), m_MouseMoveEndColor, m_Percent
                    Else
                         DrawLonghorn m_MidColor, RGB(81, 107, 148), m_EndColor, m_Percent
                    End If
                    If (hasFocus = True) Or ((Ambient.DisplayAsDefault = True) And (showFocusR = True)) Then
                        DrawFocusR
                    End If
             Case 5
                  If disyellowrect = True Then
                     DrawRectangle 0, 0, Wi, He, m_OfficeXpMousemoveFrameColor, True
                     DrawRectangle 1, 1, Wi - 2, He - 2, m_OfficeXpMousemoveFillColor
                  Else
                     DrawRectangle 0, 0, Wi, He, m_OfficeXpFrameColor, True
                     DrawRectangle 1, 1, Wi - 2, He - 2, m_OfficeXpFillColor
                  End If
                  If hasFocus = True Then DrawFocusR
        End Select
        SetTextColor .hdc, cText 'restore font color

        If UserControl.TextWidth(m_Caption) <= Wi Then
           DrawText .hdc, m_Caption, -1, Rc, DT_CENTERABS
        Else
           NumLine = (UserControl.TextWidth(m_Caption) \ Wi) + 1
           Rc.Top = (He - UserControl.TextHeight(m_Caption) * NumLine) \ 2
           DrawText .hdc, m_Caption, -1, Rc, DT_CENTER Or DT_WORDBREAK '
        End If

    ElseIf curStat = 2 Then
'#@#@#@#@#@# 按钮按下 #@#@#@#@#@#
        Select Case MyButtonType

            Case 1 'Windows XP
                 Select Case MyXpType
                        Case 1
                             DrawJianBian 171, 170, 188, 1, He - 2, 1 'XP银色效果
                         Case 2
                              stepXP1 = 15 / He
                              XPface = ShiftColor(cFace, &H30, True)
                              XPface = ShiftColor(XPface, -32, True)
                              For I = 1 To He
                                  DrawLine 0, He - I, Wi, He - I, ShiftColor(XPface, -stepXP1 * I, True)
                              Next
                          Case 3
                               DrawJianBian 224, 224, 216, 1, He - 2, 1 'XP蓝色效果
                End Select
                SetTextColor UserControl.hdc, cText
                If UserControl.TextWidth(m_Caption) <= Wi Then
                   DrawText UserControl.hdc, m_Caption, -1, rc2, DT_CENTERABS
                Else
                   rc2.Top = Rc.Top + 2
                   DrawText .hdc, m_Caption, -1, rc2, DT_CENTER Or DT_WORDBREAK '
                End If
                
                DrawXPFrame &H733C00, &H7B4D10
            
                If hasFocus = True Then DrawFocusR
            Case 2 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                DrawLine 2, 0, Wi - 2, 0, cDarkShadow
                DrawLine 2, He - 1, Wi - 2, He - 1, cDarkShadow
                DrawLine 0, 2, 0, He - 2, cDarkShadow
                DrawLine Wi - 1, 2, Wi - 1, He - 2, cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                mSetPixel 2, 2, ShiftColor(cShadow, -&H40)
                mSetPixel 3, 3, ShiftColor(cShadow, -&H20)
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cShadow
                DrawLine 1, He - 3, Wi - 2, He - 3, cShadow
                mSetPixel Wi - 4, He - 4, cShadow
                DrawLine Wi - 2, 3, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                DrawLine 3, He - 2, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 2, He - 3, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 3, He - 2, ShiftColor(cShadow, -&H20)

                mSetPixel 2, He - 2, ShiftColor(cShadow, -&H20)
                mSetPixel 2, He - 3, ShiftColor(cShadow, -&H10)
                mSetPixel 1, He - 3, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 2, 2, ShiftColor(cShadow, -&H20)
                mSetPixel Wi - 3, 2, ShiftColor(cShadow, -&H10)
                mSetPixel Wi - 3, 1, ShiftColor(cShadow, -&H10)
                SetTextColor .hdc, cLight
                If UserControl.TextWidth(m_Caption) <= Wi Then
                   DrawText UserControl.hdc, m_Caption, -1, rc2, DT_CENTERABS
                Else
                   rc2.Top = Rc.Top + 2
                   DrawText .hdc, m_Caption, -1, rc2, DT_CENTER Or DT_WORDBREAK '
                End If
                'DrawText .hdc, m_Caption, -1, rc2, DT_CENTERABS
                
             Case 3
                      If MyColorType = 1 Then
                          DrawMacOS m_MouseDownMidColor, RGB(203, 130, 62), m_MouseDownEndColor, m_Percent
                      Else
                          DrawMacOS m_MouseDownMidColor, RGB(84, 157, 84), m_MouseDownEndColor, m_Percent, True
                      End If
                      SetTextColor .hdc, cText
                      If .TextWidth(m_Caption) <= Wi Then
                         DrawText .hdc, m_Caption, -1, rc2, DT_CENTERABS
                      Else
                         rc2.Top = Rc.Top + 2
                         DrawText .hdc, m_Caption, -1, rc2, DT_CENTER Or DT_WORDBREAK '
                      End If
                      If hasFocus = True Then DrawFocusR
               Case 4
                       DrawLonghorn m_MouseDownMidColor, RGB(81, 107, 148), m_MouseDownEndColor, m_Percent
                       SetTextColor .hdc, cText
                       If .TextWidth(m_Caption) <= Wi Then
                          DrawText .hdc, m_Caption, -1, rc2, DT_CENTERABS
                       Else
                          rc2.Top = Rc.Top + 2
                          DrawText .hdc, m_Caption, -1, rc2, DT_CENTER Or DT_WORDBREAK '
                       End If
                        If hasFocus = True Then DrawFocusR
             Case 5
                DrawRectangle 0, 0, Wi, He, m_OfficeXpMousemoveFrameColor, True
                If MyColorType = 3 Then
                   DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColorCustom(m_OfficeXpMousemoveFillColor, -40, -30, -15, RGB(250, 250, 250))
                Else
                   DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColorCustom(m_OfficeXpMousemoveFillColor, -30, -30, -60, RGB(250, 250, 250))
                End If
                If .TextWidth(m_Caption) <= Wi Then
                   DrawText .hdc, m_Caption, -1, rc2, DT_CENTERABS
                Else
                   rc2.Top = Rc.Top + 2
                   DrawText .hdc, m_Caption, -1, rc2, DT_CENTER Or DT_WORDBREAK '
                End If
                If hasFocus = True Then DrawFocusR
        End Select
    End If
Else
'#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
    Select Case MyButtonType

        Case 1 'Windows XP
            DrawRectangle 0, 0, Wi, He, RGB(244, 244, 234)
            SetTextColor UserControl.hdc, RGB(161, 162, 146)
            If .TextWidth(m_Caption) <= Wi Then
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTERABS
            Else
               NumLine = (.TextWidth(m_Caption) \ Wi) + 1
               Rc.Top = (He - .TextHeight(m_Caption) * NumLine) \ 2
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTER Or DT_WORDBREAK '
            End If
            DrawXPFrame RGB(201, 199, 186), RGB(201, 199, 186)
          
        Case 2 'Mac
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight
            SetTextColor .hdc, cShadow
            DrawText .hdc, m_Caption, -1, Rc, DT_CENTERABS
            DrawLine 2, 0, Wi - 2, 0, cDarkShadow
            DrawLine 2, He - 1, Wi - 2, He - 1, cDarkShadow
            DrawLine 0, 2, 0, He - 2, cDarkShadow
            DrawLine Wi - 1, 2, Wi - 1, He - 2, cDarkShadow
            mSetPixel 1, 1, cDarkShadow
            mSetPixel 1, He - 2, cDarkShadow
            mSetPixel Wi - 2, 1, cDarkShadow
            mSetPixel Wi - 2, He - 2, cDarkShadow
            mSetPixel 1, 2, cFace
            mSetPixel 2, 1, cFace
            DrawLine 3, 2, Wi - 3, 2, cHighLight
            DrawLine 2, 2, 2, He - 3, cHighLight
            mSetPixel 3, 3, cHighLight
            DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
            DrawLine 1, He - 3, Wi - 3, He - 3, cFace
            mSetPixel Wi - 4, He - 4, cFace
            DrawLine Wi - 2, 3, Wi - 2, He - 2, cShadow
            DrawLine 3, He - 2, Wi - 2, He - 2, cShadow
            mSetPixel Wi - 3, He - 3, cShadow
            mSetPixel 2, He - 2, cFace
            mSetPixel 2, He - 3, cLight
            mSetPixel Wi - 2, 2, cFace
            mSetPixel Wi - 3, 2, cLight
            SetTextColor .hdc, cHighLight
            If .TextWidth(m_Caption) <= Wi Then
               DrawText .hdc, m_Caption, -1, rc2, DT_CENTERABS
            Else
               NumLine = (.TextWidth(m_Caption) \ Wi) + 1
               rc2.Top = (He - .TextHeight(m_Caption) * NumLine) \ 2 + 2
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTER Or DT_WORDBREAK '
            End If
        Case 3
            DrawRectangle 1, 1, Wi - 2, He - 2, RGB(241, 242, 237)
            DrawXPFrame RGB(204, 204, 202), RGB(204, 204, 202)
            SetTextColor .hdc, RGB(171, 168, 153)
            If .TextWidth(m_Caption) <= Wi Then
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTERABS
            Else
               NumLine = (.TextWidth(m_Caption) \ Wi) + 1
               Rc.Top = (He - .TextHeight(m_Caption) * NumLine) \ 2
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTER Or DT_WORDBREAK '
            End If

        Case 4
            DrawLonghorn RGB(198, 218, 255), RGB(128, 140, 162), RGB(232, 242, 255), 53
            SetTextColor .hdc, RGB(171, 168, 153)
            If .TextWidth(m_Caption) <= Wi Then
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTERABS
            Else
               NumLine = (.TextWidth(m_Caption) \ Wi) + 1
               Rc.Top = (He - .TextHeight(m_Caption) * NumLine) \ 2
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTER Or DT_WORDBREAK '
            End If
        Case 5
            If MyColorType = [Force Standard] Then
               DrawRectangle 0, 0, Wi, He, m_OfficeXpFrameColor, True
               DrawRectangle 1, 1, Wi - 2, He - 2, m_OfficeXpFillColor
               SetTextColor .hdc, m_OfficeXpFrameColor
            Else
               DrawRectangle 0, 0, Wi, He, RGB(188, 188, 64), True
               DrawRectangle 1, 1, Wi - 2, He - 2, RGB(238, 239, 208)
               SetTextColor .hdc, RGB(188, 188, 64)

            End If
            If .TextWidth(m_Caption) <= Wi Then
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTERABS
            Else
               NumLine = (.TextWidth(m_Caption) \ Wi) + 1
               Rc.Top = (He - .TextHeight(m_Caption) * NumLine) \ 2
               DrawText .hdc, m_Caption, -1, Rc, DT_CENTER Or DT_WORDBREAK '
            End If
    End Select
End If
End With
'restore focus value
hasFocus = preFocusValue

End Sub

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
'this is my custom function to draw rectangles and frames
'it's faster and smoother than using the line method

Dim bRect As RECT
Dim hBrush As Long
Dim Ret As Long

bRect.Left = x
bRect.Top = y
bRect.Right = x + Width
bRect.Bottom = y + Height

hBrush = CreateSolidBrush(Color)

If OnlyBorder = False Then
    Ret = FillRect(UserControl.hdc, bRect, hBrush)
Else
    Ret = FrameRect(UserControl.hdc, bRect, hBrush)
End If

Ret = DeleteObject(hBrush)
End Sub

Private Sub DrawLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
'a fast way to draw lines
Dim pt As POINTAPI

UserControl.ForeColor = Color
MoveToEx UserControl.hdc, X1, Y1, pt
LineTo UserControl.hdc, X2, Y2

End Sub
Private Sub DrawJianBian(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal Lower As Long, ByVal Uper As Long, ByVal Steper As Long)
Dim I As Integer
Dim rval As Single
Dim gval As Single
Dim bval As Single
   rval = R
   gval = G
   bval = B
   For I = Lower To Uper Step Steper
          rval = rval + (255 - R) / (He - 5)
          gval = gval + (255 - G) / (He - 5)                        '画XP渐变色
          bval = bval + (255 - B) / (He - 5)
          If rval > 255 Then rval = 255
          If gval > 255 Then gval = 255
          If bval > 255 Then bval = 255
          DrawLine 1, I, Wi - 1, I, RGB(rval, gval, bval)
   Next
End Sub
Private Sub DrawXPFrame(ByVal FrameColor As Long, ByVal PointColor As Long)
  
  DrawLine 2, 0, Wi - 2, 0, FrameColor
  DrawLine 2, He - 1, Wi - 2, He - 1, FrameColor
  DrawLine 0, 2, 0, He - 2, FrameColor                      '画XP外框和字
  DrawLine Wi - 1, 2, Wi - 1, He - 2, FrameColor
  mSetPixel 1, 1, PointColor
  mSetPixel 1, He - 2, PointColor
  mSetPixel Wi - 2, 1, PointColor
  mSetPixel Wi - 2, He - 2, PointColor
End Sub
Private Sub DrawMacOS(ByVal SaturatedColor As Long, ByVal FrameColor As Long, ByVal TingeColor As Long, ByVal mPercent As Integer, Optional isFill As Boolean = False)
  '画Mac OS 形状
  Dim I As Integer
  Dim RedColor As Long, BlueColor As Long, GreenColor As Long
  Dim RedColor1 As Long, BlueColor1 As Long, GreenColor1 As Long
  Dim rval As Single, gval As Single, bval As Single
  Dim Rcl As Single, Bcl As Single, Gcl As Single
  Dim rval1 As Single, gval1 As Single, bval1 As Single
  
    BlueColor1 = ((TingeColor \ &H10000) Mod &H100)
    GreenColor1 = ((TingeColor \ &H100) Mod &H100)
    RedColor1 = (TingeColor And &HFF)
    BlueColor = ((SaturatedColor \ &H10000) Mod &H100)
    GreenColor = ((SaturatedColor \ &H100) Mod &H100)
    RedColor = (SaturatedColor And &HFF)
    If isFill = False Then
       rval = RedColor
       gval = GreenColor
       bval = BlueColor
             Bcl = (BlueColor1 - BlueColor) / (He * mPercent / 100 - 1)
             Gcl = (GreenColor1 - GreenColor) / (He * mPercent / 100 - 1)
             Rcl = (RedColor1 - RedColor) / (He * mPercent / 100 - 1)
       For I = He * mPercent / 100 To 1 Step -1
             DrawLine 1, I, Wi - 1, I, RGB(rval, gval, bval)
             rval = rval + Rcl
             gval = gval + Gcl
             bval = bval + Bcl
             If rval > 255 Then rval = 255
             If rval < 0 Then rval = 0
             If gval > 255 Then gval = 255
             If gval < 0 Then gval = 0
             If bval > 255 Then bval = 255
             If bval < 0 Then bval = 0
        Next
          rval = rval - 2 * Rcl: rval1 = rval - 3 * Rcl
        gval = gval - 2 * Gcl: gval1 = gval - 3 * Gcl
        bval = bval - 2 * Bcl: bval1 = bval - 3 * Bcl
        For I = 0 To Wi - 4
             rval = 255 - I * (255 - RedColor1) \ (Wi - 10)
             gval = 255 - I * (255 - GreenColor1) \ (Wi - 10)         '画一条渐变亮色线
             bval = 255 - I * (255 - BlueColor1) \ (Wi - 10)
             If rval < 0 Then rval = 0
             If gval < 0 Then gval = 0
             If bval < 0 Then bval = 0
             mSetPixel I + 2, 1, RGB(rval, gval, bval)
        Next
        For I = 0 To Wi - Wi \ 3 - 2
             rval1 = 255 - I * (255 - RedColor1) \ (Wi - 2 * Wi \ 3 - 2)
             gval1 = 255 - I * (255 - GreenColor1) \ (Wi - 2 * Wi \ 3 - 2)        '画一条亮色线
             bval1 = 255 - I * (255 - BlueColor1) \ (Wi - 2 * Wi \ 3 - 2)
             If rval1 < 0 Then rval1 = 0
             If gval1 < 0 Then gval1 = 0
             If bval1 < 0 Then bval1 = 0
             mSetPixel I + 2, 2, RGB(rval1, gval1, bval1)
        Next
             Bcl = (BlueColor1 - BlueColor) / (He - He * mPercent / 100 - 4)      '
             Gcl = (GreenColor1 - GreenColor) / (He - He * mPercent / 100 - 4)    ''此段让颜色均匀变化
             Rcl = (RedColor1 - RedColor) / (He - He * mPercent / 100 - 4)
        DrawLine 1, He * mPercent / 100 + 1, Wi - 1, He * mPercent / 100 + 1, SaturatedColor
        DrawLine 1, He * mPercent / 100 + 2, Wi - 1, He * mPercent / 100 + 2, SaturatedColor
        rval = RedColor
        gval = GreenColor
        bval = BlueColor
        For I = He * mPercent / 100 + 3 To He - 4
               DrawLine 1, I, Wi - 1, I, RGB(rval, gval, bval)
               rval = rval + Rcl
               gval = gval + Gcl
               bval = bval + Bcl
               If rval > 255 Then rval = 255
               If rval < 0 Then rval = 0
               If gval > 255 Then gval = 255
               If gval < 0 Then gval = 0
               If bval > 255 Then bval = 255
               If bval < 0 Then bval = 0
        Next
        If He < 10 Then Exit Sub
        Bcl = GetPixel(UserControl.hdc, 2, He - 10)                 '借用Bcl为变量画线
        If Bcl < 0 Then Bcl = 0
        DrawLine 1, He - 9, Wi - 1, He - 9, Bcl
        DrawLine 1, He - 8, Wi - 1, He - 8, Bcl
        DrawLine 1, He * mPercent / 100 + 5, Wi - 1, He * mPercent / 100 + 5, SaturatedColor
        DrawLine 1, He - 3, Wi - 2, He - 3, Bcl
        DrawLine 1, He - 2, Wi - 2, He - 2, Bcl
   Else
       DrawRectangle 1, 1, Wi - 2, He - 2, SaturatedColor
   End If
   DrawXPFrame FrameColor, FrameColor
End Sub


Private Sub mSetPixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
    Call SetPixelV(UserControl.hdc, x, y, Color)
End Sub
'**************************************************************************
Private Sub DrawFocusR()                                           '画焦点框
    SetTextColor UserControl.hdc, cText
    DrawFocusRect UserControl.hdc, rc3
End Sub
Private Sub SetColors()

If MyColorType = 2 Then              'Custom
    cFace = BackC
    cText = ForeC
    cShadow = ShiftColor(cFace, -&H40)
    cLight = ShiftColor(cFace, &H1F)
    cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
    cDarkShadow = ShiftColor(cFace, -&HC0)
    'Select Case MyButtonType
    '       Case 3, 4
    '            m_MidColor = RGB(197, 197, 197)
    '            m_EndColor = RGB(238, 238, 238)
    '            m_MouseMoveMidColor = RGB(123, 167, 211)
    '            m_MouseMoveEndColor = RGB(210, 245, 251)
    '            m_MouseDownMidColor = RGB(227, 167, 105)
    '            m_MouseDownEndColor = RGB(252, 220, 199)
    '
    'End Select
ElseIf MyColorType = 3 Then              'ForceStandard
    cFace = &HC0C0C0
    cShadow = &H808080
    cLight = &HDFDFDF
    cDarkShadow = &H0
    cHighLight = &HFFFFFF
    cText = &H0
    m_OfficeXpFillColor = RGB(236, 233, 216)
    m_OfficeXpFrameColor = RGB(172, 169, 154)
    m_OfficeXpMousemoveFillColor = RGB(193, 210, 238)
    m_OfficeXpMousemoveFrameColor = RGB(49, 105, 198)
Else
'if MyColorType is 1 or has not been set then use windows colors
    cFace = GetSysColor(COLOR_BTNFACE)
    cShadow = GetSysColor(COLOR_BTNSHADOW)
    cLight = GetSysColor(COLOR_BTNLIGHT)
    cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
    cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
    cText = ForeC                                              'GetSysColor(COLOR_BTNTEXT)
    Select Case MyButtonType
           Case 4
                m_MidColor = RGB(197, 197, 197)
                m_EndColor = RGB(238, 238, 238)
                m_MouseMoveMidColor = RGB(123, 167, 211)
                m_MouseMoveEndColor = RGB(210, 245, 251)
                m_MouseDownMidColor = RGB(227, 167, 105)
                m_MouseDownEndColor = RGB(252, 220, 199)
           'Case 9
                
    End Select
End If
End Sub

Private Sub MakeRegion()
'this function creates the regions to "cut" the UserControl
'so it will be transparent in certain areas

Dim rgn1 As Long, rgn2 As Long
    
    DeleteObject rgnNorm
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    
Select Case MyButtonType

    Case 3, 4, 8, 9 'Windows XP and Mac
        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1

End Select

DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
'设置访问键

Dim ampersandPos As Long

If Len(m_Caption) > 1 Then
    ampersandPos = InStr(1, m_Caption, "&", vbTextCompare)
    If (ampersandPos < Len(m_Caption)) And (ampersandPos > 0) Then
        If Mid(m_Caption, ampersandPos + 1, 1) <> "&" Then 'if text is sonething like && then no access key should be assigned, so continue searching
            UserControl.AccessKeys = LCase(Mid(m_Caption, ampersandPos + 1, 1))
        Else 'do only a second pass to find another ampersand character
            ampersandPos = InStr(ampersandPos + 2, m_Caption, "&", vbTextCompare)
            If Mid(m_Caption, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid(m_Caption, ampersandPos + 1, 1))
            Else
                UserControl.AccessKeys = ""
            End If
        End If
    Else
        UserControl.AccessKeys = ""
    End If
Else
    UserControl.AccessKeys = ""
End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal value As Long, Optional isXP As Boolean = False) As Long
'this function will add or remove a certain color
'quantity and return the result

Dim Red As Long, Blue As Long, Green As Long

If isXP = False Then
    Blue = ((Color \ &H10000) Mod &H100) + value
Else
    Blue = ((Color \ &H10000) Mod &H100)
    Blue = Blue + ((Blue * value) \ &HC0)
End If
Green = ((Color \ &H100) Mod &H100) + value
Red = (Color And &HFF) + value
    
    'check red
    If Red < 0 Then
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    'check green
    If Green < 0 Then
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    'check blue
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If

ShiftColor = RGB(Red, Green, Blue)
End Function
Private Function ShiftColorCustom(ByVal Color As Long, ByVal Rcl As Long, ByVal Gcl As Long, ByVal Bcl As Long, ByVal Color1 As Long) As Long
Dim Red As Long, Blue As Long, Green As Long
Dim Red1 As Long, Blue1 As Long, Green1 As Long
'If isOnly = True Then
   Blue = ((Color \ &H10000) Mod &H100) + Bcl
   Green = ((Color \ &H100) Mod &H100) + Gcl
   Red = (Color And &HFF) + Rcl
   Blue1 = ((Color1 \ &H10000) Mod &H100)
   Green1 = ((Color1 \ &H100) Mod &H100)
   Red1 = (Color1 And &HFF)
    'check red
    If Red < 0 Then
        Red = 0
    ElseIf Red > Red1 Then
        Red = Red1
    End If
    'check green
    If Green < 0 Then
        Green = 0
    ElseIf Green > Green1 Then
        Green = Green1
    End If
    'check blue
    If Blue < 0 Then
        Blue = 0
    ElseIf Blue > Blue1 Then
        Blue = Blue1
    End If
'Else
    
'End If
ShiftColorCustom = RGB(Red, Green, Blue)
End Function
Private Sub DrawLonghorn(ByVal SaturatedColor As Long, ByVal FrameColor As Long, ByVal TingeColor As Long, ByVal mPercent As Integer)
  '画Longhorn 形状
  Dim I As Integer
  Dim RedColor As Long, BlueColor As Long, GreenColor As Long
  Dim RedColor1 As Long, BlueColor1 As Long, GreenColor1 As Long
  Dim rval As Long, gval As Long, bval As Long
  Dim Rcl As Long, Bcl As Long, Gcl As Long
  
    BlueColor1 = ((TingeColor \ &H10000) Mod &H100)
    GreenColor1 = ((TingeColor \ &H100) Mod &H100)
    RedColor1 = (TingeColor And &HFF)
    BlueColor = ((SaturatedColor \ &H10000) Mod &H100)
    GreenColor = ((SaturatedColor \ &H100) Mod &H100)
    RedColor = (SaturatedColor And &HFF)
       rval = RedColor + 10
       gval = GreenColor + 10
       bval = BlueColor + 10
             Bcl = (BlueColor1 - BlueColor) / (He * mPercent / 100 - 1)
             Gcl = (GreenColor1 - GreenColor) / (He * mPercent / 100 - 1) '此段让颜色均匀变化值
             Rcl = (RedColor1 - RedColor) / (He * mPercent / 100 - 1)
       For I = He * mPercent / 100 To 3 Step -1
             DrawLine 1, I, Wi - 1, I, RGB(rval, gval, bval)
             rval = rval + Rcl
             gval = gval + Gcl
             bval = bval + Bcl
             If rval > 255 Then rval = 255
             If gval > 255 Then gval = 255
             If bval > 255 Then bval = 255
             If rval < 0 Then rval = 0
             If gval < 0 Then gval = 0
             If bval < 0 Then bval = 0
        Next
        
             Bcl = (BlueColor1 - BlueColor) / (He - He * mPercent / 100 - 4)      '
             Gcl = (GreenColor1 - GreenColor) / (He - He * mPercent / 100 - 4)    ''此段让颜色均匀变化值
             Rcl = (RedColor1 - RedColor) / (He - He * mPercent / 100 - 4)
        For I = He * mPercent / 100 + 1 To He * mPercent / 100 + He * 0.27
              DrawLine 1, I, Wi - 1, I, SaturatedColor                            '连续画深色线
        Next
        rval = RedColor + 5
        gval = GreenColor + 5
        bval = BlueColor + 5
        For I = He * (mPercent + 27) / 100 To He - 4
               DrawLine 1, I, Wi - 1, I, RGB(rval, gval, bval)
               rval = rval + Rcl
               gval = gval + Gcl
               bval = bval + Bcl
               If rval > 255 Then rval = 255
               If gval > 255 Then gval = 255
               If bval > 255 Then bval = 255
               If rval < 0 Then rval = 0
               If gval < 0 Then gval = 0
               If bval < 0 Then bval = 0
        Next
        DrawRectangle 1, 1, Wi - 2, He - 2, RGB(211, 229, 255), True
        DrawRectangle 2, 2, Wi - 4, He - 4, RGB(201, 216, 255), True
        DrawLine 1, He - 3, Wi - 1, He - 3, RGB(207, 234, 255)
        DrawXPFrame FrameColor, FrameColor
End Sub


'**************************************************************************
'     timer事件处理鼠标移出
'＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊＊
Private Sub Timer1_Timer()
    Dim pnt As POINTAPI
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt

    If pnt.x < UserControl.ScaleLeft Or _
       pnt.y < UserControl.ScaleTop Or _
       pnt.x > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
       
        Timer1.Enabled = False
        RaiseEvent MouseOut
        disyellowrect = False
        Call Redraw(0, True)
    End If
End Sub

