Attribute VB_Name = "ModuleIniFileOperation"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 读写ini文件
'
'****************************************************************
Option Explicit

'API声明
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'****************************************************************
'
' Ini文件写操作
' 参数：
' sFile - Ini文件名
' sSection - 段
' sKeyName - 键名
' sKeyValue - 写入键值
' 返回值：
' 成功返回True，失败返回False
'
'****************************************************************
Public Function WriteIni(ByVal sFile As String, ByVal sSection As String, ByVal sKeyName As String, ByVal sKeyValue As String) As Boolean
    
    'API执行返回值
    Dim lRet As Long
      
    '写操作
    lRet = WritePrivateProfileString(sSection, sKeyName, sKeyValue, sFile)
    
    If lRet = 0 Then
        '写文件失败，返回
        WriteIni = False
        Exit Function
    Else
        '成功，返回
        WriteIni = True
        Exit Function
    End If
    
End Function

'****************************************************************
'
' Ini文件读操作
' 参数：
' sFile - Ini文件名
' sSection - 段
' sKeyName - 键值
' 返回值：
' 读取到的字符串
'
'****************************************************************
Public Function ReadIni(ByVal sFile As String, ByVal sSection As String, ByVal sKeyName As String) As String

    'API返回值
    Dim lRet As Long

    Dim sTemp As String * 255
       
    '读操作
    lRet = GetPrivateProfileString(sSection, sKeyName, "", sTemp, 255, sFile)
    
    If lRet = 0 Then
        '返回空
        ReadIni = ""
        Exit Function
    Else
        '去掉不可见字符
        ReadIni = TrimNull(sTemp)
    End If
    
End Function

