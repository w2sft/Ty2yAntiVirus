Attribute VB_Name = "ModuleProcessPriority"
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

'打开进程常量
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_SET_INFORMATION As Long = &H200

'正常优先级=32，低于正常优先级=16384
Public Function SetProcessPriority(ByVal Priority As Long) As Boolean
    
    Dim lProcessID As Long
    '取得当前进程ID
    lProcessID = GetCurrentProcessId()
    
    
    '打开进程
    Dim lOpenProcessReturn As Long
    lOpenProcessReturn = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_SET_INFORMATION, 0, lProcessID)
    
    If lOpenProcessReturn Then
        '设置优先级
        Call SetPriorityClass(lOpenProcessReturn, Priority)
        '设置函数返回值
        SetProcessPriority = True
    Else
        '设置函数返回值
        SetProcessPriority = False
    End If
   
    '关闭句柄
    Call CloseHandle(lOpenProcessReturn)
End Function

