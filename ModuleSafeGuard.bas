Attribute VB_Name = "ModuleSafeGuard"
'****************************************************************
'
' 自我保护
'
'****************************************************************
Option Explicit

'API声明
Private Declare Function SafeGuardON Lib "SafeGuard.dll" (ByVal TargetHwnd As Long) As Boolean
Private Declare Function SafeGuardOFF Lib "SafeGuard.dll" () As Boolean
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'****************************************************************
'
' 开启或关闭自我保护，防止进程被非法结束
' 参数：Boolean型，传入True时启动自我保护，传入False时关闭自我保护
' 返回值：True表示成功，False表示失败
'
'****************************************************************
Public Function SafeGuard(ByVal bOperation As Boolean) As Boolean

    Dim lHwnd As Long
    '取得当前进程Pid值
    lHwnd = GetCurrentProcessId
    
    Dim bRet As Boolean
    '判断操作行为，执行保护或停止保护
    If bOperation = True Then
    
        bRet = SafeGuardON(lHwnd)
        '返回结果
        '注：VC中return true返回的true值，跟VB中true值不同。故以下代码不能使用bret=true进行比较。
        'VC True=1
        'VB true=0ffffff
        If bRet Then
            SafeGuard = True
        Else
            SafeGuard = False
        End If
    Else
        SafeGuardOFF
        SafeGuard = True
    End If
    
End Function

