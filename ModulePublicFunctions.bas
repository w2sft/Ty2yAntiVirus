Attribute VB_Name = "ModulePublicFunctions"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 共公函数
'
'****************************************************************

Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function HookMsgOFF Lib "MsgCallBack.dll" () As Boolean
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'****************************************************************
'
' 扩展功能的Trim，可去掉字符串中的chr(0)，即不可见字符，回车\换行
' 参数：字符串
' 返回值：去掉chr(0)后的字符串
'
'****************************************************************
Public Function TrimNull(sStr As String) As String
    If InStr(1, sStr, Chr(0)) Then
        TrimNull = Trim(Left(sStr, InStr(1, sStr, Chr(0)) - 1))
    Else
        TrimNull = Trim(sStr)
    End If
End Function

'****************************************************************
'
' 退出程序
' 恢复子类化、停止Hook，完整退出程序
'
'****************************************************************
Public Function TrueEnd()
  
    '恢复窗体消息处理
    SetWindowLong FormScan.hWnd, -4, lSubClassOldAddress
    
    '移除托盘图标
    RemoveTrayIcon
    
    '改变窗体卸载标志
    bEnableUnloadForm = True
    '卸载窗体
    
    Unload FormSetting
    Unload FormUpdate
    Unload FormSelectFolders
    Unload FormMenu
    Unload FormScan
    Unload FormSkin
    Unload FormAlertMessage
    Unload FormScanFinish
    Unload FormAlertVirus
    Unload FormAbout
    
End Function

'****************************************************************
'
' 使软件使用最少内存
'
'****************************************************************
Public Function MiniUseMemory()
    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
End Function

'****************************************************************
'
'产生随机数标题
'
'****************************************************************
Public Function RndChr(Numbers As Long) As String
    Dim I As Long
    Dim sAscii As String
    For I = 1 To Numbers
        '对随机数生成器做初始化的动作。
        Randomize
        '生成65-90之间的随机数值。
        sAscii = Int((90 - 65) * Rnd + 65)
        RndChr = RndChr + Chr(sAscii)
    Next I
    RndChr = UCase(RndChr)
End Function

Public Function ReSkinAll()
    FormAlertMessage.ReSkinMe
    DoEvents
    FormAlertVirus.ReSkinMe
    DoEvents
    FormScan.ReSkinMe
    DoEvents
    FormScanFinish.ReSkinMe
    DoEvents
    FormSelectFolders.ReSkinMe
    DoEvents
    FormSetting.ReSkinMe
    DoEvents
    FormSkin.ReSkinMe
    DoEvents
    FormUpdate.ReSkinMe
    DoEvents
    FormAbout.ReSkinMe
    DoEvents
End Function
