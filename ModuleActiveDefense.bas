Attribute VB_Name = "ModuleActiveDefense"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 主动防御控制
'
'****************************************************************
Option Explicit

'API声明
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWow64Process As Boolean) As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
               
Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Const PROCESSOR_ARCHITECTURE_AMD64 As Long = &H9
Private blnIsWin64bit As Boolean

Private Function APIFunctionPresent(ByVal FunctionName As String, ByVal DLLName As String) As Boolean
    Dim lHandle As Long
    Dim lAddr  As Long
    Dim FreeLib As Boolean
    FreeLib = False
    lHandle = GetModuleHandle(DLLName)
    If lHandle = 0 Then
        lHandle = LoadLibraryEx(DLLName, 0&, &H1)
        FreeLib = True
    End If
    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        If FreeLib Then
            FreeLibrary lHandle
        End If
    End If
    APIFunctionPresent = (lAddr <> 0)
End Function

'****************************************************************
'
' 判断系统是32位还是64位
' 字符串，含64或32信息
'
'****************************************************************
Private Function InfoVersion64bit() As String
    Dim lngRet As Long
    Dim strTemp As String
    Dim Si As SYSTEM_INFO
    blnIsWin64bit = False

    If APIFunctionPresent("GetNativeSystemInfo", "kernel32") Then
        GetNativeSystemInfo Si
        If Si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 Then
                strTemp = "64bits"
                blnIsWin64bit = True
        Else
                strTemp = "32bits"
                blnIsWin64bit = False
        End If
    End If
    InfoVersion64bit = strTemp
End Function
    
'****************************************************************
'
' 开启病毒防护
'
'****************************************************************
Public Function OnTimeProtectON()
    '判断是否是64位系统
    Call InfoVersion64bit
    
    If blnIsWin64bit = True Then
        If Right(App.Path, 1) = "\" Then
            Call ShellExecute(FormScan.hWnd, "open", App.Path & "OnTimeProtectService64.exe", 0, 0, 1)
        Else
            Call ShellExecute(FormScan.hWnd, "open", App.Path & "\OnTimeProtectService64.exe", 0, 0, 1)
        End If
    Else
        If Right(App.Path, 1) = "\" Then
            Call ShellExecute(FormScan.hWnd, "open", App.Path & "OnTimeProtectService32.exe", 0, 0, 1)
        Else
            Call ShellExecute(FormScan.hWnd, "open", App.Path & "\OnTimeProtectService32.exe", 0, 0, 1)
        End If
    End If
End Function





