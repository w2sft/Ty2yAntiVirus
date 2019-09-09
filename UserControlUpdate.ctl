VERSION 5.00
Begin VB.UserControl UserControlUpdate 
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2370
   ScaleHeight     =   1755
   ScaleWidth      =   2370
   Begin VB.Image ImageLogo 
      Height          =   630
      Left            =   0
      Picture         =   "UserControlUpdate.ctx":0000
      Top             =   0
      Width           =   630
   End
End
Attribute VB_Name = "UserControlUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 升级软件和病毒库控件
'
'****************************************************************

'API声明
Option Explicit
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal lFile As Long, bBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpQueryInfo Lib "WinInet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal bBuffer As Any, ByRef lsBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetSetFilePointer Lib "WinInet.dll" (ByVal lFile As Long, ByVal lDistanceToMove As Long, ByVal pReserved As Long, ByVal dwMoveMethod As Long, ByVal dwContext As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'常量定义
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const scUserAgent = "Wing"
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000

'控件内公供变量
Private sUrl As String
Private sSaveFile As String

Private bConnect As Boolean
Private lOpen As Long
Private lFile As Long
Private sBuffer As String
Private lBuffer As Long
Private bRetQueryInfo As Boolean


'控件公供事件，可在外部调用

'下载进度
Public Event DownloadProcess(lProgress As Long)
'下载完成事件
Public Event DownLoadFinish()
'错误信息
Public Event ErrorMassage(sDescription As String)

'****************************************************************
'
' 初始化
'
'****************************************************************
Public Function Init() As Boolean
    
    bConnect = True
    '该函数是第一个由应用程序调用的 WinINet 函数。它告诉 Internet DLL 初始化内部数据结构并准备接收应用程序之后的其他调用。
    '当应用程序结束使用 Internet 函数时，应调用 InternetCloseHandle 函数来释放与之相关的资源
    lOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    
    '设置bConnect状态可以中止连接
    If bConnect = False Then
        CancelDownload
        
        '用户中止，函数返回失败
        Init = False
        Exit Function
    End If
    
    If lOpen = 0 Then
        CancelDownload
        '创建网络连接失败
        bConnect = False
        
        '函数返回失败
        Init = False
        
        '引发错误事件
        RaiseEvent ErrorMassage("创建网络连接失败。")
        Exit Function
    Else
    
        '通过一个完整的FTP，Gopher或HTTP网址打开一个资源。
        lFile = InternetOpenUrl(lOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        
        If bConnect = False Then
            CancelDownload
            
            '用户中止，函数返回失败
            Init = False
            Exit Function
        End If
        
        If lFile = 0 Then
            CancelDownload
            '无法连接到服务器。
            bConnect = False
            
            '函数返回失败
            Init = False
            
            '引发错误事件
            RaiseEvent ErrorMassage("无法连接到升级服务器。")
            Exit Function
        Else
            sBuffer = Space(1024)
            lBuffer = 1024
            
            '从服务器获取 HTTP 请求头信息
            bRetQueryInfo = HttpQueryInfo(lFile, 21, sBuffer, lBuffer, 0)
            
            '获取成功
            If bRetQueryInfo Then
                
                sBuffer = Mid(sBuffer, 1, lBuffer)
                '函数返回成功
                Init = True
                Exit Function
            Else
                sBuffer = ""
                '函数返回失败
                Init = False
                Exit Function
            End If
        End If
    End If

End Function

'****************************************************************
'
' 开始升级
'
'****************************************************************
Public Function StartUpdate() As Boolean
    
    '定义错误出口
    'On Error GoTo ErrorExit
    
    Dim bBuffer(1 To 1024) As Byte
    Dim lRet As Long
        
    '已下载大小
    Dim lDownloadSize As Long
    Dim I As Long
    
    '用户中止
    If bConnect = False Then
        CancelDownload
        StartUpdate = False
        Exit Function
    End If
    
    '打开文件序号，如：#1
    Dim lFileNumber As Long
    lFileNumber = FreeFile()
    
    '下载保存文件方式:下载保存为file.exe.bak 下载完后 删除file.exe 重命名 file.exe.bak 为file.exe
    Dim sTempDownloadFile As String
    sTempDownloadFile = Trim(sSaveFile) & ".bak"
    
    '如果文件已经存在，先删除
    If Dir(sTempDownloadFile) <> "" Then
        Call DeleteFile(sTempDownloadFile)
        DoEvents
    End If
    
    '打开文件，保存下载数据
    Open sTempDownloadFile For Binary Access Write As #lFileNumber
        
        Do
            '读取被下载文件内容
            InternetReadFile lFile, bBuffer(1), 1024, lRet
            
            DoEvents
            '完整块1024
            If lRet = 1024 Then
                If bConnect = False Then
                    StartUpdate = False
                    Close #lFileNumber
                    GoTo UserCancel
                End If
                
                '写入文件
                Put #1, , bBuffer
            Else
                '剩余部分块
                For I = 1 To lRet
                    Put #1, , bBuffer(I)
                    DoEvents
                Next I
            End If
            
            '已经下载的字节数
            lDownloadSize = lDownloadSize + lRet
            
            '引发下载进度事件
            RaiseEvent DownloadProcess(lDownloadSize)
            
        Loop Until lRet < 1024
    Close #lFileNumber
    
    '检查已下载字节数与要下载文件大小是否一至，一至说明下载完成了。
    If lDownloadSize = CLng(GetHeader("Content-Length")) Then
        
        '下载完成，删除原文件，重命名下载文件为原文件名
        If Dir(Trim(sSaveFile)) <> "" Then
        
            If Dir(Trim(sSaveFile)) <> "" Then
                Call DeleteFile(sSaveFile)
            End If
            DoEvents
                
        End If
        
        Sleep 50
        DoEvents
        
        If Dir(sSaveFile) = "" Then
            '重命名下载的文件
            Name sTempDownloadFile As sSaveFile
        End If
        DoEvents
        
    End If

    '下载完成,引发下载完成事件
    RaiseEvent DownLoadFinish
    
'用户操作，正常中止
UserCancel:
    CancelDownload
    Exit Function
   
'出错退出
ErrorExit:
    '引发错误事件
    RaiseEvent ErrorMassage("升级过程发生意外错误。")
    CancelDownload
    Close #lFileNumber
    
End Function

'****************************************************************
'
' 用户中止升级过程
'
'****************************************************************
Public Function CancelDownload()
    bConnect = False
    InternetCloseHandle lOpen
    InternetCloseHandle lFile
End Function


'****************************************************************
'
' 获取要下载文件的信息
'
'****************************************************************
Public Function GetHeader(Optional hdrName As String) As String
    Dim sTemp As Long
    Dim sTemp2 As String
    
    '用户中止
    If bConnect = False Then
        GetHeader = "0"
        CancelDownload
        Exit Function
    End If
    
    If sBuffer <> "" Then
        Select Case UCase(hdrName)
        '文件大小
        Case "CONTENT-LENGTH"
            sTemp = InStr(sBuffer, "Content-Length")
            sTemp2 = Mid(sBuffer, sTemp + 16, Len(sBuffer))
            sTemp = InStr(sTemp2, Chr(0))
            GetHeader = CStr(Mid(sTemp2, 1, sTemp - 1))
            
        Case "CONTENT-TYPE"
            sTemp = InStr(sBuffer, "Content-Type")
            sTemp2 = Mid(sBuffer, sTemp + 14, Len(sBuffer))
            sTemp = InStr(sTemp2, Chr(0))
            GetHeader = CStr(Mid(sTemp2, 1, sTemp - 1))
        
        Case "DATE"
            sTemp = InStr(sBuffer, "Date")
            sTemp2 = Mid(sBuffer, sTemp + 6, Len(sBuffer))
            sTemp = InStr(sTemp2, Chr(0))
            GetHeader = CStr(Mid(sTemp2, 1, sTemp - 1))
        '最后修改时间
        Case "LAST-MODIFIED"
            sTemp = InStr(sBuffer, "Last-Modified")
            sTemp2 = Mid(sBuffer, sTemp + 15, Len(sBuffer))
            sTemp = InStr(sTemp2, Chr(0))
            GetHeader = CStr(Mid(sTemp2, 1, sTemp - 1))
            
        Case "SERVER"
            sTemp = InStr(sBuffer, "Server")
            sTemp2 = Mid(sBuffer, sTemp + 8, Len(sBuffer))
            sTemp = InStr(sTemp2, Chr(0))
            GetHeader = CStr(Mid(sTemp2, 1, sTemp - 1))
            
        Case vbNullString
            GetHeader = sBuffer
            Case Else
            GetHeader = "0"
            
        End Select
    Else
        GetHeader = "0"
    End If

End Function

'****************************************************************
'
' 控件设置属性,保存文件名
'
'****************************************************************
Public Property Let SaveFile(ByVal sInFileName As String)
    sSaveFile = sInFileName
End Property

'****************************************************************
'
' 控件设置属性，URL地址
'
'****************************************************************
Public Property Let URL(ByVal sInUrl As String)
    sUrl = sInUrl
End Property

Private Sub UserControl_Resize()
    With UserControl
        .Width = ImageLogo.Width
        .Height = ImageLogo.Height
    End With
End Sub


