VERSION 5.00
Begin VB.Form FormUpdate 
   BorderStyle     =   0  'None
   Caption         =   "升级"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   8940
   Icon            =   "FormUpdate.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormUpdate.frx":57E2
   ScaleHeight     =   6030
   ScaleWidth      =   8940
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Index           =   3
      Left            =   7320
      ScaleHeight     =   2535
      ScaleWidth      =   135
      TabIndex        =   12
      Top             =   1440
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   2
      Left            =   1080
      ScaleHeight     =   135
      ScaleWidth      =   6375
      TabIndex        =   11
      Top             =   3720
      Width           =   6375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Index           =   1
      Left            =   968
      ScaleHeight     =   2535
      ScaleWidth      =   135
      TabIndex        =   10
      Top             =   1320
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Index           =   0
      Left            =   1080
      ScaleHeight     =   135
      ScaleWidth      =   6375
      TabIndex        =   9
      Top             =   1328
      Width           =   6375
   End
   Begin VB.ListBox ListUpdate 
      Height          =   2400
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   6375
   End
   Begin Ty2yAntiVirus.Command CommandStop 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "停止更新"
   End
   Begin Ty2yAntiVirus.Command CommandUpdate 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "检测更新"
   End
   Begin VB.Timer TimerAutoUpdate 
      Interval        =   1000
      Left            =   4080
      Top             =   4560
   End
   Begin Ty2yAntiVirus.UserControlUpdate UserControlUpdate1 
      Height          =   630
      Left            =   7560
      TabIndex        =   0
      Top             =   1200
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1111
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   3135
      Left            =   840
      Top             =   1080
      Width           =   7470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "更新倒记时："
      Height          =   180
      Left            =   3720
      TabIndex        =   7
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label LabelUpdateLeftTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   90
   End
   Begin VB.Label LabelAutoUpdateMsg 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   2
      Left            =   8280
      Picture         =   "FormUpdate.frx":B4FDC
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   1
      Left            =   8280
      Picture         =   "FormUpdate.frx":B573E
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   0
      Left            =   8280
      Picture         =   "FormUpdate.frx":B5EA0
      Top             =   0
      Width           =   465
   End
   Begin VB.Label LabelAutoUpdateState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   180
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自动更新："
      Height          =   180
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "FormUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 软件、病毒库升级
' 升级方式说明：本地设置文件中，保存着整体版本号、软件版本号、特征码版本号共三个不同的版本号。
' 升级时，从网络文件中获取最新整体版本号，与当前软件中的整体版本号进行比对，如果网络中的整体
' 版本号更高，则进行更新，下载升级文件。每次升级把设置文件同时也下载，以同步版本号。(Setting.ini)
' 网络文件格式："Update"(6位，用于标识是升级校验文件)+整体版本号(6位)+待升级文件+"$"
'
'****************************************************************

Option Explicit

'api声明
Private Declare Function InternetOpen Lib "WinInet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "WinInet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "WinInet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "WinInet.dll" (ByVal hInet As Long) As Integer
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

'常量定义
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000

'是否可以下载下一个文件
Dim bEnableDownloadNext As Boolean
'是否中止下载
Dim bStopDownload As Boolean
'当前下载的文件大小
Dim lCurrentFileSize
'自动升级计数
Dim lCurrentSecond As Long

Dim lFinishedCount As Long

'****************************************************************
'
' 获取指定网址返回内容
' 参数：URL地址
' 返回值：返回字网页内容，可为空
'
'****************************************************************
Public Function DetectUpdateFile(sUrl As String) As String
    
    Dim lInternetOpenUrl As Long
    Dim lInternetOpen As Long
    Dim sTemp As String * 1024
    Dim lInternetReadFile As Long
    Dim lSize As Long
    
    Dim sContent As String
    sContent = vbNullString
          
    '初始化，以使用 WinINet 函数
    lInternetOpen = InternetOpen("WangLiwen", 1, vbNullString, vbNullString, 0)
          
    If lInternetOpen Then
        '打开网址
        lInternetOpenUrl = InternetOpenUrl(lInternetOpen, sUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
        
        If lInternetOpenUrl Then
            Do
                '读取网页内容
                lInternetReadFile = InternetReadFile(lInternetOpenUrl, sTemp, 1024, lSize)
                sContent = sContent & Mid(sTemp, 1, lSize)
            Loop While (lSize <> 0)
        Else
            '显示连接错误
            If Me.Visible = True Then
                MsgBox "错误代码：001。连接升级服务器失败，请检查网络连接状况。"
            Else
                LabelAutoUpdateMsg.Caption = "错误代码：001。连接升级服务器失败，请检查网络连接状况。"
            End If
            DetectUpdateFile = ""
            Exit Function
        End If
        
        lInternetReadFile = InternetCloseHandle(lInternetOpenUrl)
    Else
        '显示连接错误
        If Me.Visible = True Then
            MsgBox "错误代码：002。连接升级服务器失败，请检查网络连接状况。"
        Else
            LabelAutoUpdateMsg.Caption = "错误代码：002。连接升级服务器失败，请检查网络连接状况。"
        End If
        DetectUpdateFile = ""
        Exit Function
    End If
    
    '升级查询返回以“Update”开始为标识
    If UCase(Left(sContent, 6)) = UCase("Update") Then
        DetectUpdateFile = Right(sContent, Len(sContent) - 6)
    Else
        DetectUpdateFile = ""
    End If
    
End Function

'****************************************************************
'
' 检测升级更新内容
'
'****************************************************************
Private Sub CommandUpdate_Click()
    lFinishedCount = 0
    
    '是否中止下载标志
    bStopDownload = False

    '从指定网址获取持更新文件内容
    Dim sUpdateFiles As String
    sUpdateFiles = DetectUpdateFile("http://www.ty2y.com/ty2yantivirus/update/update.txt")
    
    If sUpdateFiles = "" Then
        If Me.Visible = True Then
            MsgBox "升级服务器错误，无法获取升级文件。"
        Else
            LabelAutoUpdateMsg.Caption = "升级服务器错误，无法获取升级文件。"
        End If
        Exit Sub
    End If
    
    '取得新整体版本号
    Dim lNewWholeVersion As Long
    lNewWholeVersion = Left(sUpdateFiles, 6)
    sUpdateFiles = Right(sUpdateFiles, Len(sUpdateFiles) - 6)
    
    '软件版本记录文件
    Dim sVersionFile As String
    If Right(App.Path, 1) = "\" Then
        sVersionFile = App.Path & "Version.ini"
    Else
        sVersionFile = App.Path & "\Version.ini"
    End If
        
    '读取旧整体版本
    Dim lOldWholeVersion As Long
    lOldWholeVersion = ReadIni(sVersionFile, "Version", "WholeVersion")
    
    '判断是否有升级内容
    If lNewWholeVersion <= lOldWholeVersion Then
        If Me.Visible = True Then
            MsgBox "已是最新版本，无需升级。"
        Else
            LabelAutoUpdateMsg.Caption = "已是最新版本，无需升级。"
        End If
        Exit Sub
    End If
    
    '清空旧升级内容
    ListUpdate.Clear
    
    CommandUpdate.Enabled = False
    CommandStop.Enabled = True
    
    '如果有更新内容
    Do While Not sUpdateFiles = ""
        DoEvents
        '从返回内容中取得持更新文件名
        With ListUpdate
            .AddItem Left(sUpdateFiles, InStr(1, sUpdateFiles, "$") - 1)
        End With
        sUpdateFiles = Right(sUpdateFiles, Len(sUpdateFiles) - InStr(1, sUpdateFiles, "$"))
    Loop
    
    '升级更新每个文件
    Dim I As Long
    For I = 0 To ListUpdate.ListCount - 1
        DoEvents
        '是否中止下载
        If bStopDownload = True Then
            Exit For
        End If
        
        '如果是要下载exe文件，先用movefileex去掉它的执行保护
        If InStr(1, LCase(ListUpdate.List(I)), LCase(".exe")) Then
            Call MoveFileEx(ListUpdate.List(I), RndChr(6), 1)
        End If
        
        '是否可以更新下一个文件，下载依次进行，正在下载文件时不能进行此操作
        If bEnableDownloadNext = True Then
            With UserControlUpdate1
                '文件下载地址
                .URL = "http://www.ty2y.com/ty2yantivirus/update/" & ListUpdate.List(I)
                '本地保存地址
                .SaveFile = App.Path & "\" & ListUpdate.List(I)
                If .Init = True Then
                    '当前下载的文件大小，单位：byte
                    lCurrentFileSize = .GetHeader("Content-Length")
                    '设置正在下载标志
                    bEnableDownloadNext = False
                    '开始下载文件
                    .StartUpdate
                End If
            End With
        End If
    Next
    
End Sub

Private Sub CommandStop_Click()
    bStopDownload = True
    UserControlUpdate1.CancelDownload
    CommandUpdate.Enabled = True
    CommandStop.Enabled = False
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
    If ImageExit(2).Visible = True Then
        Me.Hide
    End If
End Sub

'窗体启动函数
Private Sub Form_Load()
    ReSkinMe
    Dim I As Long
    For I = 0 To 2

        '初始化关闭铵钮位置
        With ImageExit(I)
            .Left = 8280
            .Top = 0
        End With
    Next
    
    '自动升级计数
    lCurrentSecond = 0
    
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False
    
    '设置正在下载标志
    bEnableDownloadNext = True
    
    '软件设置记录文件
    Dim sSettingsFile As String
    If Right(App.Path, 1) = "\" Then
        sSettingsFile = App.Path & "Settings.ini"
    Else
        sSettingsFile = App.Path & "\Settings.ini"
    End If
    
    '读取升级设置
    Dim lAutoUpdate As Long
    lAutoUpdate = ReadIni(sSettingsFile, "Update", "AutoUpdate")
        
    '自动更新设置
    If lAutoUpdate = 1 Then
        LabelAutoUpdateState.Caption = "已启用"
        TimerAutoUpdate.Enabled = True
    Else
        LabelAutoUpdateState.Caption = "未启用"
        TimerAutoUpdate.Enabled = False
    End If
    
    CommandUpdate.Enabled = True
    CommandStop.Enabled = False
        
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

'自动升级
Private Sub TimerAutoUpdate_Timer()
    
    Dim sSettingsFile As String
    
    '软件设置记录文件
    If Right(App.Path, 1) = "\" Then
        sSettingsFile = App.Path & "Settings.ini"
    Else
        sSettingsFile = App.Path & "\Settings.ini"
    End If
    
    '自动升级频率
    Dim lAutoUpdateIntervel As Long
    lAutoUpdateIntervel = ReadIni(sSettingsFile, "Update", "AutoUpdateIntervel")
    
    '升级频率变为秒
    Dim lMaxSecond As Long
    lMaxSecond = lAutoUpdateIntervel * 3600
    
    '自动升级计数
    lCurrentSecond = lCurrentSecond + 1
    LabelUpdateLeftTime.Caption = 3600 - lCurrentSecond & "秒"
    
    '是否到达自动升级时间
    If lCurrentSecond >= lMaxSecond Then
        '升级计数清零
        lCurrentSecond = 0
        '升级
        CommandUpdate_Click
    End If
End Sub

'****************************************************************
'
' 升级控件事件：一个下载完成
'
'****************************************************************
Private Sub UserControlUpdate1_DownLoadFinish()
    lFinishedCount = lFinishedCount + 1
    bEnableDownloadNext = True
    Dim I As Long
    'For i = 0 To ListUpdate.ListCount - 1

        Dim sListText As String
        sListText = ListUpdate.List(ListUpdate.ListCount - 1)
        
        'If InStr(sListText, "100%") Then
        
            '如果是exe文件,下载后运行
            If InStr(LCase(sListText), LCase(".exe")) Then
                        
                DoEvents
                                
                Dim sRunFile  As String
                sRunFile = sListText
                                
                '下载后执行
                Call WinExec(sRunFile, 1)
                DoEvents
            End If
                        
            If ListUpdate.ListCount = lFinishedCount Then
            
                LabelAutoUpdateMsg.Caption = ""
                
                CommandUpdate.Enabled = True
                CommandStop.Enabled = False
                If Me.Visible = True Then
                    MsgBox "升级完成。"
                Else
                    LabelAutoUpdateMsg.Caption = "升级完成。"
                    
                    '提醒窗体
                    Dim uAlertForm As New FormAlertMessage
                    
                    '弹出窗体提示用户
                    With uAlertForm
                        '设置窗体位置
                        .Top = Screen.Height - .Height - 300
                        .Left = Screen.Width - .Width
                                    
                        '升级完成提示
                        Dim sUpdated  As String
                        sUpdated = "自动升级完成，已升级文件：" & vbCrLf
                        Dim j As Long
                        For j = 0 To ListUpdate.ListCount - 1
                            sUpdated = sUpdated & ListUpdate.List(j) & " "
                        Next j
                        
                        '显示内容
                        .LabelInfo.Caption = sUpdated
                        '显示窗体
                        .Show vbModal
                    End With
                    
                End If
                
                '刷新软件版本、病毒库版本、Web控件
                FormScan.TimerUpdateRefresh.Enabled = True
            End If
        'End If
    'Next
End Sub

'****************************************************************
'
' 升级控件事件：实时下载进展
'
'****************************************************************
Private Sub UserControlUpdate1_DownloadProcess(lProgress As Long)

    LabelAutoUpdateMsg.Caption = "正在下载，请稍后…… " & ((lProgress * 100) \ lCurrentFileSize) & "%"
    Exit Sub
    
    bEnableDownloadNext = False
    Dim I As Long
    For I = 0 To ListUpdate.ListCount - 1
        Dim sListText As String
        sListText = ListUpdate.List(I)
        
        If InStr(sListText, "%") = 0 Then
            sListText = sListText & " " & ((lProgress * 100) \ lCurrentFileSize) & "%"
        Else
            sListText = Left(sListText, InStr(sListText, " ") - 1)
            sListText = sListText & " " & ((lProgress * 100) \ lCurrentFileSize) & "%"
        End If
        
        If InStr(ListUpdate.List(I), "100%") = 0 Then
            ListUpdate.List(I) = sListText
        End If
        
    Next
End Sub

'****************************************************************
'
' 升级控件事件：错误消息
'
'****************************************************************
Private Sub UserControlUpdate1_ErrorMassage(sDescription As String)
    If Me.Visible = True Then
        MsgBox sDescription
    Else
        LabelAutoUpdateMsg.Caption = sDescription
    End If
End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Update2.bmp")
        .ImageExit(0).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit0.bmp")
        .ImageExit(1).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit1.bmp")
        .ImageExit(2).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit2.bmp")
    End With
End Function
