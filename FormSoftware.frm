VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FormSoftware 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   Picture         =   "FormSoftware.frx":0000
   ScaleHeight     =   7680
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin SHDocVwCtl.WebBrowser WebBrowserSoftware 
      Height          =   6980
      Left            =   210
      TabIndex        =   0
      Top             =   505
      Width           =   9400
      ExtentX         =   16581
      ExtentY         =   12312
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   2
      Left            =   9240
      Picture         =   "FormSoftware.frx":F4844
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   1
      Left            =   9240
      Picture         =   "FormSoftware.frx":F4FA6
      Top             =   0
      Width           =   465
   End
   Begin VB.Image ImageExit 
      Height          =   285
      Index           =   0
      Left            =   9240
      Picture         =   "FormSoftware.frx":F5708
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "FormSoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' 黑白名单管理窗体
'
'****************************************************************
Option Explicit

'api声明
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '按下鼠标左键
    If Button = vbLeftButton Then
        '为当前的应用程序释放鼠标捕获
        ReleaseCapture
        '移动窗体
        SendMessage Me.hwnd, &HA1, 2, 0
    End If
End Sub

'窗体启动函数
Private Sub Form_Load()
    ReSkinMe
    Dim j As Long
    For j = 0 To 2
        '初始化关闭铵钮位置
        With ImageExit(j)
            .Left = 9240
            .Top = 0
        End With
    Next
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False
    
    '加载html页
    If Right(App.Path, 1) = "\" Then
        WebBrowserSoftware.Navigate App.Path & "software.html"
    Else
        WebBrowserSoftware.Navigate App.Path & "\software.html"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '关闭铵钮
    ImageExit(0).Visible = True
    ImageExit(1).Visible = False
    ImageExit(2).Visible = False
End Sub

Private Sub ImageExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub ImageExit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '退出铵钮状态
    ImageExit(0).Visible = False
    ImageExit(1).Visible = True
    ImageExit(2).Visible = False
End Sub

Private Sub ImageExit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '退出铵钮状态
    ImageExit(0).Visible = False
    ImageExit(1).Visible = False
    ImageExit(2).Visible = True
End Sub

Private Sub ImageExit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    '退出点击铵钮
    Unload Me
End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\software.bmp")
        .ImageExit(0).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit0.bmp")
        .ImageExit(1).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit1.bmp")
        .ImageExit(2).Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\Exit2.bmp")
    End With
End Function


