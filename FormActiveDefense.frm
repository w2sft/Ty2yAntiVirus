VERSION 5.00
Begin VB.Form FormActiveDefense 
   BorderStyle     =   0  'None
   Caption         =   "系统防护"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   Icon            =   "FormActiveDefense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormActiveDefense.frx":57E2
   ScaleHeight     =   3180
   ScaleWidth      =   5760
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerBringToTop 
      Interval        =   10
      Left            =   5160
      Top             =   1800
   End
   Begin Ty2yAntiVirus.Command CommandRun 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      font            =   "FormActiveDefense.frx":41224
      xptype          =   3
      caption         =   "运行"
   End
   Begin Ty2yAntiVirus.Command CommandBlock 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      font            =   "FormActiveDefense.frx":41248
      xptype          =   3
      caption         =   "禁止"
   End
   Begin VB.CheckBox CheckRemember 
      BackColor       =   &H00FFEFDE&
      Caption         =   "记住本操作"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Image ImageFileIconShow 
      Height          =   480
      Left            =   360
      Top             =   720
      Width           =   480
   End
   Begin VB.Label LabelHowToDoAction 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应该运行还是禁止?"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   180
      Left            =   1080
      MouseIcon       =   "FormActiveDefense.frx":4126C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2160
      Width           =   1530
   End
   Begin VB.Image ImageDefault16x16Icon 
      Height          =   240
      Left            =   360
      Picture         =   "FormActiveDefense.frx":413BE
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "是否允许启动?"
      Height          =   180
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   1170
   End
   Begin VB.Image ImageFileIcon 
      Height          =   480
      Left            =   360
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LabelCommand 
      Caption         =   "Command"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   4080
      Width           =   5175
   End
   Begin VB.Label LabelSignature 
      Caption         =   "Signature"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   5175
   End
   Begin VB.Label LabelFile 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "FormActiveDefense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' 检测到未知文件启动时的拦截窗体
'
'****************************************************************

Option Explicit

'API声明
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'点击运行铵钮
Private Sub CommandRun_Click()
    
    '把铵钮不可操作化
    CommandRun.Enabled = False
    CommandBlock.Enabled = False
    
    '判断是否取得了文件图片
    If ImageFileIcon = 0 Then
        ImageFileIcon.Picture = ImageDefault16x16Icon.Picture
    End If

    '设置允许标志
    lUserAction = 1
    
    DoEvents
    '添加记录
    With FormScan
        .TextShieldLog.Text = .TextShieldLog.Text & Date & " " & Time & " " & LabelFile.Caption & " " & "放行" & " " & "手动操作"
    End With
        
    Unload Me
End Sub

'点击阻止铵钮
Private Sub CommandBlock_Click()
    
    '把铵钮不可操作化
    CommandRun.Enabled = False
    CommandBlock.Enabled = False
    
    '判断是否取得了文件图片
    If ImageFileIcon = 0 Then
        ImageFileIcon.Picture = ImageDefault16x16Icon.Picture
    End If
    
    '设置阻止标志
    lUserAction = 2
    
    '添加记录
    With FormScan
        .TextShieldLog.Text = .TextShieldLog.Text & Date & " " & Time & " " & LabelFile.Caption & " " & "阻止" & " " & "手动操作"
    End With
    
    Unload Me
End Sub

'激活窗体时
Private Sub Form_Activate()
    SetForegroundWindow Me.hWnd
End Sub

'窗体启动函数
Private Sub Form_Load()
    SetForegroundWindow Me.hWnd
    ReSkinMe
End Sub

Private Sub LabelHowToDoAction_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.haiqi.cn/", 0, 0, 1)
End Sub

'用定时器控制让窗体置于最前
Private Sub TimerBringToTop_Timer()
    
    '将窗体置于最前
    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
    DoEvents
    
    '窗体提示音
    If Dir(App.Path & "\alert.wav") <> "" Then
        PlaySound App.Path & "\alert.wav", 1, 1
    End If
    
    '恢复窗体位置
    SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
    
    TimerBringToTop.Enabled = False
End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\ActiveDefense.bmp")
    End With
End Function
