VERSION 5.00
Begin VB.Form FormSkin 
   BorderStyle     =   0  'None
   Caption         =   "请选择皮肤"
   ClientHeight    =   2970
   ClientLeft      =   4560
   ClientTop       =   0
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   Picture         =   "FormSkin.frx":0000
   ScaleHeight     =   2970
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox ComboSkin 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin Ty2yAntiVirus.Command CommandOK 
      Height          =   375
      Left            =   3480
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "皮肤列表："
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   2
      Left            =   4560
      Picture         =   "FormSkin.frx":31842
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   1
      Left            =   4560
      Picture         =   "FormSkin.frx":31FA4
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   0
      Left            =   4560
      Picture         =   "FormSkin.frx":32706
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "FormSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 皮肤选择
'
'****************************************************************
Option Explicit

'api声明
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

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
        cFileName As String * 1024
        cAlternate As String * 256
End Type

Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Sub CommandOK_Click()

    '改变皮肤
    sSkin = ComboSkin.Text
    '载入皮肤
    ReSkinAll
    DoEvents
    
    Dim sSettingsFile As String
    
    '软件设置记录文件
    If Right(App.Path, 1) = "\" Then
        sSettingsFile = App.Path & "Settings.ini"
    Else
        sSettingsFile = App.Path & "\Settings.ini"
    End If
    
    '写入配置文件
    Call WriteIni(sSettingsFile, "Normal", "Skin", sSkin)
    
    Unload Me
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

'窗体启动函数
Private Sub Form_Load()
    ReSkinMe
    Dim j As Long
    For j = 0 To 2
        '初始化关闭铵钮位置
        With ImageExit(j)
            .Left = 4560
            .Top = 0
        End With
    Next
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False
    
    '初始化皮肤列表
    ComboSkin.Clear
    
    Dim sSkinPath As String
    sSkinPath = App.Path
    If Right(sSkinPath, 1) <> "\" Then
        sSkinPath = sSkinPath & "\"
    End If
    sSkinPath = sSkinPath & "Skin\"
    InitSkins sSkinPath, "*.*"
    
    '当前皮肤
    ComboSkin.Text = sSkin
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

'----------------------------------------------------------------#
'
' 函数名：InitSkins
' 功能：查找皮肤
' 返回值：无
'
'----------------------------------------------------------------#
Public Sub InitSkins(DirPath As String, FileSpec As String)
    'API用自定义结构。
    Dim FindData As WIN32_FIND_DATA
    
    '要搜索的目录
    DirPath = Trim(DirPath)
    '构成完整目录形式
    If Right(DirPath, 1) <> "\" Then
        DirPath = DirPath & "\"
    End If
    
    'FindFirstfile返回的句柄
    Dim FindHandle As Long
    
    '在目标目录中取得第一个文件名
    FindHandle = FindFirstFile(DirPath & FileSpec, FindData)
    
    '如果没有失败(说明有文件)
    If FindHandle <> 0 Then
        If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
      
             '如果是一个目录
            If Left(FindData.cFileName, 1) <> "." And Left(FindData.cFileName, 2) <> ".." Then
                
                '添加到目录列中
                ComboSkin.AddItem Trim(FindData.cFileName)
                
            End If
            
        End If
    End If
    
    '现在开始找其它文件
    If FindHandle <> 0 Then
        
        Dim sFullName As String
        
        'FindNextFile返回的句柄
        Dim FindNextHandle As Long
        
        Do
           
            DoEvents
                            
            '找下一个文件
            FindNextHandle = FindNextFile(FindHandle, FindData)
            If FindNextHandle <> 0 Then
                    
                If FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                        
                    '是目录的话,就加到目录列表
                    If Left(FindData.cFileName, 1) <> "." And Left(FindData.cFileName, 2) <> ".." Then
                        
                        ComboSkin.AddItem Trim(FindData.cFileName)
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    End If
    
    '关闭句柄
    Call FindClose(FindHandle)
    
End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Skin.bmp")
        .ImageExit(0).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit0.bmp")
        .ImageExit(1).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit1.bmp")
        .ImageExit(2).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit2.bmp")
    End With
End Function
