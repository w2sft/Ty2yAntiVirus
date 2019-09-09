VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormRightClickScanResult 
   BorderStyle     =   0  'None
   Caption         =   "右键扫描结果"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   Icon            =   "FormRightClickScanResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormRightClickScanResult.frx":57E2
   ScaleHeight     =   7680
   ScaleWidth      =   9780
   StartUpPosition =   3  '窗口缺省
   Begin Ty2yAntiVirus.Command CommandUnload 
      Height          =   375
      Left            =   8040
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpType          =   3
      Caption         =   "确定"
   End
   Begin MSComctlLib.ListView ListViewScanResult 
      Height          =   6150
      Left            =   190
      TabIndex        =   0
      Top             =   510
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   10848
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "文件"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "扫描结果"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Timer TimerBringToTop 
      Interval        =   10
      Left            =   2040
      Top             =   0
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   0
      Left            =   9120
      Picture         =   "FormRightClickScanResult.frx":FA024
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   1
      Left            =   9120
      Picture         =   "FormRightClickScanResult.frx":FA786
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   2
      Left            =   9120
      Picture         =   "FormRightClickScanResult.frx":FAEE8
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "FormRightClickScanResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
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

Private Sub CommandUnload_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    ReSkinMe
    SetForegroundWindow Me.hWnd
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
        
    DoEvents
    Dim i As Long
    For i = 0 To 2
        '初始化关闭铵钮位置
        With ImageExit(i)
            .Left = 9120
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
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\RightClickScanResult.bmp")
        .ImageExit(0).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit0.bmp")
        .ImageExit(1).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit1.bmp")
        .ImageExit(2).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit2.bmp")
    End With
End Function
