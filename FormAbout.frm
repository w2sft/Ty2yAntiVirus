VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   0  'None
   Caption         =   "关于"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormAbout.frx":0000
   ScaleHeight     =   3180
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin Ty2yAntiVirus.Command CommandOK 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
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
      Caption         =   "确定"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.7"
      Height          =   180
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   540
   End
   Begin VB.Label LabelHomePage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.ty2y.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ty2y杀毒软件"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 关于窗口
'
'****************************************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CommandOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ReSkinMe
    LabelHomePage.ForeColor = vbBlue
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LabelHomePage.ForeColor = vbBlue
    LabelHomePage.Font.Underline = False
End Sub

Private Sub LabelHomePage_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.ty2y.com/", 0, 0, 1)
End Sub

Private Sub LabelHomePage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LabelHomePage.ForeColor = vbRed
    LabelHomePage.Font.Underline = True
End Sub

Public Function ReSkinMe()
    With Me
        .Picture = LoadPicture(App.Path & "\Skin\" & sSkin & "\About.bmp")
    End With
End Function

