Attribute VB_Name = "ModuleSubClass"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 子类化相关模块
'
'****************************************************************
Option Explicit

'API声明
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal lParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'自定义类型
Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3

Private Const WM_SYSCOMMAND = &H112
Private Const SC_RESTORE = &HF120&
'Explorer出错时，重建托盘图标
Public WM_TASKBARCREATED   As Long

'****************************************************************
'
' 主动防御消息处理函数
' 接收Dll拦截到的进程启动消息，并进行处理
' DLL函数根据此操作结果进行放行或禁止运行操作
' 参数：句柄、自定义的消息类型，被创建进程父进程PID，进程路径等
'
'****************************************************************
Public Function SubClassMessage(ByVal lHwnd As Long, ByVal lMessage As Long, ByVal lParentPID As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    
    Dim sTemp As String
    Dim uCDS As COPYDATASTRUCT
    
    '双击弹出窗口
    If lParam = &H203 Then
        
        '扫描窗体
        With FormScan
            .Top = (Screen.Height - .Height) / 2
            .Left = (Screen.Width - .Width) / 2
            .Show
            Call SendMessage(.hWnd, WM_SYSCOMMAND, SC_RESTORE, 0)
        End With
        
        Exit Function
    End If
    
    '右键弹出菜单
    If lParam = &H204 Then
        
        SetForegroundWindow FormMenu.hWnd
        FormMenu.PopupMenu FormMenu.m_TrayMenu
        
        Exit Function
    End If
    
    '判是否是DLL传来的消息
    If lMessage = &H4A Then
        
        '使软件使用最少内存
        MiniUseMemory
    
        '取得传来的数据长度
        CopyMemory uCDS, ByVal lParam, Len(uCDS)
        sTemp = Space(uCDS.cbData)
                
        '取得传来的数据
        CopyMemory ByVal sTemp, ByVal uCDS.lpData, uCDS.cbData
        
        '数据中传来的是启动的程序
        Dim sFile As String
        sFile = sTemp
                        
        '如果文件名为空，则放行
        If TrimNull(sFile) = "" Then
            '自动放行
            SubClassMessage = 915
            Exit Function
            
        End If
        
        '如果文件不存在，则放行
        If Len(Dir(sFile)) = 0 Then
            Exit Function
        End If

        '判断是否是程序自己，则放行
        Dim sTrimFile As String
        sTrimFile = TrimNull(sFile)
        If InStr(1, LCase(sTrimFile), LCase(App.Path & "\" & App.EXEName)) Or InStr(1, LCase(App.Path & "\" & App.EXEName), LCase(sTrimFile)) Then
            '自动放行
            SubClassMessage = 915
            Exit Function
        End If
        
        If FormSetting.OptionHighShildLevel.value = True Then
            '添加记录
            With FormScan
                    .TextShieldLog.Text = .TextShieldLog.Text & Date & " " & Time & " " & sFile & " " & "阻止" & vbCrLf
            End With
                                    
            SetForegroundWindow FormScan.hWnd
            
            '弹出窗体提示用户
            Dim uAlertFormHighShield As New FormAlertMessage
            With uAlertFormHighShield
                '设置窗体位置
                .Top = Screen.Height - .Height - 300
                .Left = Screen.Width - .Width
                  
                '显示内容
                .LabelInfo.Caption = "防护等级：严格。禁止" & sFile & "启动。"
                '显示窗体
                .Show
            End With
                                   
            '退出函数
            Exit Function
        End If
        
        '判断文件大小，如果文件太大，则直接放行
        Dim lFileSize As Long
        Dim uSize As LARGE_INTEGER
        '打开文件
        lFileSize = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
        '获取文件大小
        GetFileSizeEx lFileSize, uSize
        '关闭文件
        CloseHandle lFileSize
        If uSize.lowpart / 1024 / 1024 > 15 Then
            '自动放行
            SubClassMessage = 915
            '添加记录
            With FormScan
                .TextShieldLog.Text = .TextShieldLog.Text & Date & " " & Time & " " & sFile & " " & "放行" & " " & "安全文件" & vbCrLf
            End With
            '退出函数
            Exit Function
        End If
                
        '扫描文件，判断是否是病毒
        Dim sScanResult As String
        sScanResult = ScanFile(sFile)
        If sScanResult <> "" Then
            '添加记录
            With FormScan
                    .TextShieldLog.Text = .TextShieldLog.Text & Date & " " & Time & " " & sFile & " " & "阻止" & vbCrLf
            End With
                                    
            SetForegroundWindow FormScan.hWnd
            
            '弹出窗体提示用户
            Dim uAlertForm As New FormAlertMessage
            With uAlertForm
                '设置窗体位置
                .Top = Screen.Height - .Height - 300
                .Left = Screen.Width - .Width
                  
                '显示内容
                .LabelInfo.Caption = "发现病毒，" & sFile & "(" & sScanResult & ")" & "，已拦截。"
                '显示窗体
                .Show
            End With
                                   
            '退出函数
            Exit Function
        End If
        
        '自动放行
        SubClassMessage = 915
        '添加记录
        With FormScan
            .TextShieldLog.Text = .TextShieldLog.Text & Date & " " & Time & " " & sFile & " " & "放行" & " " & "未检测到病毒" & vbCrLf
        End With
            
        '退出函数
        Exit Function
    End If
    
    '重建托盘图标
    If lMessage = WM_TASKBARCREATED Then
        AddTrayIcon
    End If
    
    '非DLL传来的数据，执行正常的消息处理函数
    SubClassMessage = CallWindowProc(lSubClassOldAddress, lHwnd, lMessage, lParentPID, lParam)
End Function

