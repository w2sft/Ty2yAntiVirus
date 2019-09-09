Attribute VB_Name = "ModuleLoadSignatures"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 校验和加载病毒库
'
'****************************************************************

Option Explicit
'****************************************************************
'
' 自定义病毒库格式
' 特征码(32位hash值) + 病毒名(32位，) + 回车换行(chr(13)chr(10))
'
'****************************************************************
Private Type Signature
    'Section hash
    sHash As String * 32
    '名称
    sName As String * 32
    '回车换行
    sVbCrlf As String * 2
End Type


'****************************************************************
'
' 公共特征码变量
' 病毒库文件由0~9、A~F 16个文件组成，每个变量名存储一个病毒库文件
'
'****************************************************************

Public uSignature0() As Signature
Public uSignature1() As Signature
Public uSignature2() As Signature
Public uSignature3() As Signature
Public uSignature4() As Signature
Public uSignature5() As Signature
Public uSignature6() As Signature
Public uSignature7() As Signature
Public uSignature8() As Signature
Public uSignature9() As Signature

Public uSignatureA() As Signature
Public uSignatureB() As Signature
Public uSignatureC() As Signature
Public uSignatureD() As Signature
Public uSignatureE() As Signature
Public uSignatureF() As Signature
Public uSignatureG() As Signature


'****************************************************************
'
' 检查病毒库文件完整性
' 返回值：病毒库完整返回 True，否则返回 False
'
'****************************************************************
Public Function CheckSignatureFile() As Boolean

    '软件当前路径
    Dim sAppPath As String
    sAppPath = App.Path
    
    '检查路径完整性
    If Right(sAppPath, 1) <> "\" Then
        sAppPath = sAppPath & "\"
    End If
    
    Dim sSignatureFile As String
    
    '检查0~9病毒库文件是否存在
    Dim I As Long
    For I = 0 To 9
        '病毒库文件
        sSignatureFile = sAppPath & I & ".sig"
        If Dir(sSignatureFile) = "" Then
            CheckSignatureFile = False
            Exit Function
        End If
    Next
    
    '检查A~Z病毒库文件是否存在
    Dim j As Long
    For j = 65 To 71
        '病毒库文件
        sSignatureFile = sAppPath & Chr(j) & ".sig"
        If Dir(sSignatureFile) = "" Then
            CheckSignatureFile = False
            Exit Function
        End If
    Next
    
    CheckSignatureFile = True
    
End Function

'****************************************************************
'
' 从病毒库文件加载特征码
'
'****************************************************************
Public Function LoadSignatures() As Long
    
    '软件当前路径
    Dim sAppPath As String
    sAppPath = App.Path
    
    '检查路径完整性
    If Right(sAppPath, 1) <> "\" Then
        sAppPath = sAppPath & "\"
    End If
    
    '病毒库文件
    Dim sSignatureFile As String
    '文件句柄
    Dim lFile As Long
    '文件长度
    Dim lFileLen As Long
    '特征码个数
    Dim lSignatureCount As Long
    
    lSignatureCount = 0
    
    '加载0~9病毒库文件
    Dim I As Long
    For I = 0 To 9
        DoEvents
        '病毒库文件
        sSignatureFile = sAppPath & I & ".sig"
        
        '从病毒库文件加载特征码
        Select Case I
        Case "0"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature0(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature0
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "1"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature1(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature1
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "2"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature2(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature2
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "3"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature3(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature3
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "4"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature4(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature4
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "5"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature5(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature5
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "6"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature6(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature6
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "7"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature7(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature7
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "8"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature8(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature8
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "9"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignature9(1 To lFileLen / 66) As Signature
                Get lFile, , uSignature9
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        End Select
        
    Next
    
    '加载A~F病毒库文件
    Dim j As Long
    For j = 65 To 70
    
        DoEvents
        '病毒库文件
        sSignatureFile = sAppPath & Chr(j) & ".sig"
        '从病毒库文件加载特征码
        Select Case UCase(Chr(j))
        Case "A"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignatureA(1 To lFileLen / 66) As Signature
                Get lFile, , uSignatureA
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "B"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignatureB(1 To lFileLen / 66) As Signature
                Get lFile, , uSignatureB
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "C"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignatureC(1 To lFileLen / 66) As Signature
                Get lFile, , uSignatureC
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "D"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignatureD(1 To lFileLen / 66) As Signature
                Get lFile, , uSignatureD
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "E"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignatureE(1 To lFileLen / 66) As Signature
                Get lFile, , uSignatureE
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        Case "F"
            lFile = FreeFile
            Open sSignatureFile For Binary As #lFile
            lFileLen = LOF(lFile)
            '判断文件是否为空
            If lFileLen <> 0 Then
                ReDim uSignatureF(1 To lFileLen / 66) As Signature
                Get lFile, , uSignatureF
                lSignatureCount = lSignatureCount + (lFileLen / 66)
            End If
            Close lFile
        End Select
    Next
    
    '"G"
    lFile = FreeFile
    Open sAppPath & "G.sig" For Binary As #lFile
    lFileLen = LOF(lFile)
    
    '判断文件是否为空
    If lFileLen <> 0 Then
        ReDim uSignatureG(1 To lFileLen / 66) As Signature
        Get lFile, , uSignatureG
        lSignatureCount = lSignatureCount + (lFileLen / 66)
    End If
    Close lFile

    
    '返回特征码总加载数
    LoadSignatures = lSignatureCount
    
    
End Function
