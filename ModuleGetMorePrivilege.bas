Attribute VB_Name = "ModuleGetMorePrivilege"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 为程序提权
'
'****************************************************************

Option Explicit

'api声明
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByVal PreviousState As Any, ReturnLength As Long) As Long

'常量定义
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Private Const SE_DEBUG_NAME = "SeDebugPrivilege"
Private Const SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege"
Private Const SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege"
Private Const SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege"
Private Const SE_INCREASE_QUOTA_NAME = ("SeIncreas=otaPrivilege")
Private Const SE_UNSOLICITED_INPUT_NAME = ("SeUnsolicitedInputPrivilege")
Private Const SE_MACHINE_ACCOUNT_NAME = ("SeMachineAccountPrivilege")
Private Const SE_TCB_NAME = ("SeTcbPrivilege")
Private Const SE_SECURITY_NAME = ("SeSecurityPrivilege")
Private Const SE_TAKE_OWNERSHIP_NAME = ("SeTakeOwnershipPrivilege")
Private Const SE_LOAD_DRIVER_NAME = ("SeLoadDriverPrivilege")
Private Const SE_SYSTEM_PROFILE_NAME = ("SeSystemProfilePrivilege")
Private Const SE_SYSTEMTIME_NAME = ("SeSystemtimePrivilege")
Private Const SE_PROF_SINGLE_PROCESS_NAME = ("SeProfileSingleProcessPrivilege")
Private Const SE_INC_BASE_PRIORITY_NAME = ("SeIncreaseBasePriorityPrivilege")
Private Const SE_CREATE_PAGEFILE_NAME = ("SeCreatePagefilePrivilege")
Private Const SE_CREATE_PERMANENT_NAME = ("SeCreatePermanentPrivilege")
Private Const SE_BACKUP_NAME = ("SeBackupPrivilege")
Private Const SE_RESTORE_NAME = ("SeRestorePrivilege")
Private Const SE_AUDIT_NAME = ("SeAuditPrivilege")
Private Const SE_SYSTEM_ENVIRONMENT_NAME = ("SeSystemEnvironmentPrivilege")
Private Const SE_CHANGE_NOTIFY_NAME = ("SeChangeNotifyPrivilege")
Private Const SE_REMOTE_SHUTDOWN_NAME = ("SeRemoteShutdownPrivilege")

'自定义类型
Private Type LUID
     lowpart As Long
     highpart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
     pLuid As LUID
     Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
     PrivilegeCount As Long ''权限的个数n
     Privileges(0) As LUID_AND_ATTRIBUTES ''如果要多个权限，数组改为(n-1)个元素
End Type

'取得调试权限,
'取得这个权限后，可以结束几乎所有的进程
'如果要在NT里关机，则需要取SE_SHUTDOWN_NAME这个权限
Public Function GetMorePrivilege()
    Dim tLuid As LUID
    Dim tp As TOKEN_PRIVILEGES, tpOld As TOKEN_PRIVILEGES
    Dim hToken As Long
    OpenProcessToken GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken
    LookupPrivilegeValue vbNullString, SE_DEBUG_NAME, tp.Privileges(0).pLuid
    tp.PrivilegeCount = 1
    tp.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    Call AdjustTokenPrivileges(hToken, False, tp, 0&, 0&, 0&)
End Function

