Attribute VB_Name = "ModuleRightClickMenu"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'建立键
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'写入启动值
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'删除项目
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'打开注册表subkey的hkey
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const REG_SZ = 1


Public Function AddFileRightClickMenu() As Boolean
    
    Dim hKey As Long, ret As Long
    
    '建立注册表项
    RegCreateKey HKEY_CLASSES_ROOT, "*\shell\海奇杀毒\command", hKey
    
    '设置右键菜单项目
    ret = RegSetValueEx(hKey, "", 0, REG_SZ, ByVal App.Path & "\" & App.EXEName & ".exe %1", ByVal LenB(StrConv(App.Path & "\" & App.EXEName & ".exe %1", vbFromUnicode)) + 1)
    
    If ret = 0 Then
    
        '成功
        AddFileRightClickMenu = True
    Else
        '失败
        AddFileRightClickMenu = False
    End If
    
    RegCloseKey hKey
    
End Function

'删除右键菜单
Public Function DeleteRightClickMenu() As Boolean

    '这里必须分步执行，如同删除文件夹一样，不能删除非空的文件夹
    DeleteKey HKEY_CLASSES_ROOT, "*\shell\海奇杀毒", "command"
    DoEvents
    DeleteKey HKEY_CLASSES_ROOT, "*\shell", "海奇杀毒"
    DoEvents
    DeleteKey HKEY_CLASSES_ROOT, "Directory\shell\海奇杀毒", "command"
    DoEvents
    DeleteKey HKEY_CLASSES_ROOT, "Directory\shell", "海奇杀毒"
    DoEvents
    DeleteRightClickMenu = True
End Function

'删除注册表键函数
Function DeleteKey(RootKey As Long, ParentKeyName As String, SubKeyName As String)

    Dim hKey As Long
    RegOpenKey RootKey, ParentKeyName, hKey
    RegDeleteKey hKey, SubKeyName
    RegCloseKey hKey
End Function


'这函数是用来单独添加到目录的
Public Function AddFolderRightClickMenu() As Boolean

    '把应用程序加入右键菜单
    Dim hKey As Long, ret As Long
    
    '建立注册表项
    RegCreateKey HKEY_CLASSES_ROOT, "Directory\shell\海奇杀毒\command", hKey
    
    '设置右键菜单项目
    ret = RegSetValueEx(hKey, "", 0, REG_SZ, ByVal App.Path & "\" & App.EXEName & ".exe %1", ByVal LenB(StrConv(App.Path & "\" & App.EXEName & ".exe %1", vbFromUnicode)) + 1)
    If ret = 0 Then
           AddFolderRightClickMenu = True
    Else
           AddFolderRightClickMenu = False
    End If
    
    RegCloseKey hKey '关闭注册表项
End Function


