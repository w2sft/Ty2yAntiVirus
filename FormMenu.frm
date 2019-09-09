VERSION 5.00
Begin VB.Form FormMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "菜单"
   ClientHeight    =   1230
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   3225
   Icon            =   "FormMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3225
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu m_TrayMenu 
      Caption         =   "托盘图标"
      Begin VB.Menu m_About 
         Caption         =   "关于"
      End
      Begin VB.Menu m_S7 
         Caption         =   "-"
      End
      Begin VB.Menu m_Homepage 
         Caption         =   "官方网站"
      End
      Begin VB.Menu m_S2 
         Caption         =   "-"
      End
      Begin VB.Menu m_Setting 
         Caption         =   "设置"
      End
      Begin VB.Menu m_S3 
         Caption         =   "-"
      End
      Begin VB.Menu m_Update 
         Caption         =   "升级"
      End
      Begin VB.Menu m_S1 
         Caption         =   "-"
      End
      Begin VB.Menu m_Exit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "FormMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 菜单
'
'****************************************************************
Option Explicit

'api声明
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Private Sub m_About_Click()
    FormAbout.Show vbModal
End Sub

Private Sub m_ClearAllShieldLog_Click()
    FormScan.TextShieldLog.Text = ""
End Sub

'退出软件
Private Sub m_Exit_Click()
    TrueEnd
End Sub

'官方网站
Private Sub m_Homepage_Click()
    Call ShellExecute(Me.hWnd, "open", "http://www.ty2y.com/", 0, 0, 1)
End Sub

'软件设置
Private Sub m_Setting_Click()
    FormSetting.Show vbModal
End Sub

'升级
Private Sub m_Update_Click()
    FormUpdate.Show vbModal
End Sub

