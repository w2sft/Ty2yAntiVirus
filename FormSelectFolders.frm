VERSION 5.00
Begin VB.Form FormSelectFolders 
   BorderStyle     =   0  'None
   Caption         =   "选择目录"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   Icon            =   "FormSelectFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormSelectFolders.frx":57E2
   ScaleHeight     =   3030
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.DriveListBox DriveScanTarget 
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
   End
   Begin Ty2yAntiVirus.Command CommandOK 
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
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
   Begin VB.Timer TimerSelectTreeviewCheck 
      Interval        =   300
      Left            =   4440
      Top             =   5160
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   0
      Left            =   4800
      Picture         =   "FormSelectFolders.frx":3A92C
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   1
      Left            =   4800
      Picture         =   "FormSelectFolders.frx":3B08E
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   2
      Left            =   4800
      Picture         =   "FormSelectFolders.frx":3B7F0
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "FormSelectFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 自定义扫描时选择扫描目标
'
'****************************************************************

Option Explicit

'api声明
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'常量定义
Private Const MAX_PATH As Long = 260
Private Const DRIVE_REMOVABLE As Long = 2
Private Const DRIVE_FIXED As Long = 3
Private Const DRIVE_REMOTE As Long = 4
Private Const DRIVE_CDROM As Long = 5
Private Const DRIVE_RAMDISK As Long = 6

Private Const TH32CS_SNAPPROCESS = &H2&
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

'自定义类型
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
        
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'点击确定铵钮
Private Sub CommandOK_Click()
    Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bEnableUnloadForm = False Then
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub ImageExit_Click(Index As Integer)
    Me.Hide
End Sub

'窗体启动函数
Private Sub Form_Load()
    
    Dim j As Long
    For j = 0 To 2
        '初始化关闭铵钮位置
        With ImageExit(j)
            .Left = 4800
            .Top = 0
        End With
    Next
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False

    DoEvents
    ReSkinMe
    
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
    
    '是否进行扫描标识
    bDoCustmerScan = False
    
End Sub

Private Sub ImageExit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '退出点击铵钮
    Me.Hide
End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\SelectFolders.bmp")
        .ImageExit(0).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit0.bmp")
        .ImageExit(1).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit1.bmp")
        .ImageExit(2).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit2.bmp")
    End With
End Function
