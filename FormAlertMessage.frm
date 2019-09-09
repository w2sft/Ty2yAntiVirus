VERSION 5.00
Begin VB.Form FormAlertMessage 
   BorderStyle     =   0  'None
   Caption         =   "提醒"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   Icon            =   "FormAlertMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormAlertMessage.frx":57E2
   ScaleHeight     =   3180
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerBringToTop 
      Interval        =   10
      Left            =   1440
      Top             =   2760
   End
   Begin VB.Timer TimerAutoUnload 
      Interval        =   1000
      Left            =   960
      Top             =   2760
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   0
      Left            =   3480
      Picture         =   "FormAlertMessage.frx":30924
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   1
      Left            =   3480
      Picture         =   "FormAlertMessage.frx":31086
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   2
      Left            =   3480
      Picture         =   "FormAlertMessage.frx":317E8
      Top             =   0
      Width           =   465
   End
   Begin VB.Label LabelLabelAutoTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "秒后自动关闭。"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label LabelAutoCloseAlertInterval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label LabelInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "info"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "FormAlertMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 信息提示窗体
'
'****************************************************************
Option Explicit

'api声明
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Sub Form_Activate()
    SetForegroundWindow Me.hWnd
    ReSkinMe
End Sub

'窗体启动函数
Private Sub Form_Load()
   
    SetForegroundWindow Me.hWnd
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '按下鼠标左键
    If Button = vbLeftButton Then
        '为当前的应用程序释放鼠标捕获
        ReleaseCapture
        '移动窗体
        SendMessage Me.hWnd, &HA1, 2, 0
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False
End Sub

Private Sub ImageExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub ImageExit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '退出铵钮状态
    ImageExit(0).Visible = False
    ImageExit(1).Visible = True
    ImageExit(2).Visible = False
End Sub

Private Sub ImageExit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '退出铵钮状态
    ImageExit(0).Visible = False
    ImageExit(1).Visible = False
    ImageExit(2).Visible = True
End Sub

Private Sub ImageExit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '退出点击铵钮
    Unload Me
End Sub

'用定时器控制让窗体自动销毁
Private Sub TimerAutoUnload_Timer()
    If LabelAutoCloseAlertInterval <> 0 Then
        LabelAutoCloseAlertInterval.Caption = CLng(LabelAutoCloseAlertInterval.Caption) - 1
        LabelAutoCloseAlertInterval.Refresh
    Else
        Unload Me
    End If
End Sub

'用定时器控制让窗体置于最前
Private Sub TimerBringToTop_Timer()
    
    '将窗体置于最前
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
    DoEvents
    
    '窗体提示音
    If Dir(App.Path & "\notify.wav") <> "" Then
        PlaySound App.Path & "\notify.wav", 1, 1
    End If
    
    '恢复窗体位置
    SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
    
    TimerBringToTop.Enabled = False

   
    '软件设置记录文件
    Dim sSettingsFile As String
    If Right(App.Path, 1) = "\" Then
        sSettingsFile = App.Path & "Settings.ini"
    Else
        sSettingsFile = App.Path & "\Settings.ini"
    End If
        
    '读取防护提醒开启状态
    Dim lAutoCloseAlert As Long
    lAutoCloseAlert = ReadIni(sSettingsFile, "Shield", "AutoCloseAlertMessage")
    
    If lAutoCloseAlert = 1 Then
    
        '自动关闭提示窗口频率
        Dim lAutoCloseAlertInterval As Long
        lAutoCloseAlertInterval = ReadIni(sSettingsFile, "Shield", "AutoCloseAlertMessageInterval")
        
        '设置时间
        LabelAutoCloseAlertInterval.Caption = lAutoCloseAlertInterval
        
        '激活自动关闭
        TimerAutoUnload = True
    Else
        TimerAutoUnload = False
        LabelAutoCloseAlertInterval.Visible = False
        LabelLabelAutoTip.Visible = False
    End If
    
    DoEvents
    Dim I As Long
    For I = 0 To 2
        '初始化关闭铵钮位置
        With ImageExit(I)
            .Left = 3480
            .Top = 0
        End With
    Next
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False

End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\ActiveDefense.bmp")
        .ImageExit(0).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit0.bmp")
        .ImageExit(1).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit1.bmp")
        .ImageExit(2).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit2.bmp")
    End With
End Function
