Attribute VB_Name = "ModuleAutorun"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 程序开机后自启动设置
'
'****************************************************************
Option Explicit

'api声明
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long                    '建立一个新的主键
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long '创建或改变一个键值,有时 lpData 应由缺省的 ByRef 型改为 ByVal 型
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'常量定义
Private Const REG_SZ = 1
Private Const HKEY_LOCAL_MACHINE = &H80000002

'设置开机启动
Public Function SetAutoRun()
    On Error Resume Next
    Dim sAddAutoRunReg As String
    sAddAutoRunReg = "HQ"
    Dim nKeyHandle As Long
    Call RegCreateKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", nKeyHandle)
    
    If Right(App.Path, 1) = "\" Then
        Call RegSetValueExString(nKeyHandle, sAddAutoRunReg, 0, REG_SZ, App.Path & App.EXEName & ".exe", 255)
    Else
        Call RegSetValueExString(nKeyHandle, sAddAutoRunReg, 0, REG_SZ, App.Path & "\" & App.EXEName & ".exe", 255)
    End If
    
    Call RegCloseKey(nKeyHandle)
End Function

'取消开机启动
Public Function StopAutoRun()
    On Error Resume Next
    Dim Ret As Long, hKey As Long
    Call RegCreateKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", hKey)
    Call RegDeleteValue(hKey, "HQ")
    Call RegCloseKey(hKey)
End Function



