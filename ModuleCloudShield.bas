Attribute VB_Name = "ModuleCloudShield"
'****************************************************************
'
' 云安全防护
'
'****************************************************************

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function InternetOpen Lib "WinInet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "WinInet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "WinInet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "WinInet.dll" (ByVal hInet As Long) As Integer
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000


'****************************************************************
'
' 获取指定网址返回值内容，用于进行云安全查询
' 参数：URL地址
' 返回值：返回字网页内容，可为空
'
'****************************************************************
Public Function CloudMatch(sUrl As String) As String
    
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
            'SetForegroundWindow FormScan.hWnd
            '显示云安全连接错误
            'MsgBox "连接云安全服务器失败，无法启用云安全防御系统，请检查网络连接状况。", vbInformation, "错误代码：001"
            CloudMatch = ""
            Exit Function
        End If
        
        lInternetReadFile = InternetCloseHandle(lInternetOpenUrl)
    Else
    
        'SetForegroundWindow FormScan.hWnd
        '显示云安全连接错误
        'MsgBox "连接云安全服务器失败，无法启用云安全防御系统，请检查网络连接状况。", vbInformation, "错误代码：002"
        CloudMatch = ""
        Exit Function
    End If
    
    '云安全防御系统查询返回：Cloud$(云安全标识)+1/0(病毒or可信文件标识)+病毒名称/可信软件名称
    If UCase(Left(sContent, 6)) = UCase("Cloud$") Then
        CloudMatch = Right(sContent, Len(sContent) - 6)
    Else
        CloudMatch = ""
    End If
    
    
End Function

