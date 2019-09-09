Attribute VB_Name = "ModuleMatchSignature"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
'特征码匹配
'成功返回病毒名称，失败返回空
'
'****************************************************************
Public Function MatchSignature(ByVal sHash As String) As String

    
    '初始化返回值
    MatchSignature = ""
    
    '检查传入的参数sHash是否为空
    If sHash = "" Then
        Exit Function
    End If
    
    '取sHash第一位为标志，确定使用特征码类
    Dim sFlagWord As String
    sFlagWord = UCase(Left(sHash, 1))
    
    Dim I As Long
    
    Select Case sFlagWord
    Case "0"
        For I = 1 To UBound(uSignature0)
            If uSignature0(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature0(I).sName
                Exit Function
            End If
        Next I
    Case "1"
        For I = 1 To UBound(uSignature1)
            If uSignature1(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature1(I).sName
                Exit Function
            End If
        Next I
    Case "2"
        For I = 1 To UBound(uSignature2)
            If uSignature2(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature2(I).sName
                Exit Function
            End If
        Next I
    Case "3"
        For I = 1 To UBound(uSignature3)
            If uSignature3(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature3(I).sName
                Exit Function
            End If
        Next I
    Case "4"
        For I = 1 To UBound(uSignature4)
            If uSignature4(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature4(I).sName
                Exit Function
            End If
        Next I
    Case "5"
        For I = 1 To UBound(uSignature5)
            If uSignature5(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature5(I).sName
                Exit Function
            End If
        Next I
    Case "6"
        For I = 1 To UBound(uSignature6)
            If uSignature6(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature6(I).sName
                Exit Function
            End If
        Next I
    Case "7"
        For I = 1 To UBound(uSignature7)
            If uSignature7(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature7(I).sName
                Exit Function
            End If
        Next I
    Case "8"
        For I = 1 To UBound(uSignature8)
            If uSignature8(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature8(I).sName
                Exit Function
            End If
        Next I
    Case "9"
        For I = 1 To UBound(uSignature9)
            If uSignature9(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignature9(I).sName
                Exit Function
            End If
        Next I
    Case "A"
        For I = 1 To UBound(uSignatureA)
            If uSignatureA(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignatureA(I).sName
                Exit Function
            End If
        Next I
    Case "B"
        For I = 1 To UBound(uSignatureB)
            If uSignatureB(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignatureB(I).sName
                Exit Function
            End If
        Next I
    Case "C"
        For I = 1 To UBound(uSignatureC)
            If uSignatureC(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignatureC(I).sName
                Exit Function
            End If
        Next I
    Case "D"
        For I = 1 To UBound(uSignatureD)
            If uSignatureD(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignatureD(I).sName
                Exit Function
            End If
        Next I
    Case "E"
        For I = 1 To UBound(uSignatureE)
            If uSignatureE(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignatureE(I).sName
                Exit Function
            End If
        Next I
    Case "F"
        For I = 1 To UBound(uSignatureF)
            If uSignatureF(I).sHash = sHash Then
                '匹配成功，返回病毒名称
                MatchSignature = uSignatureF(I).sName
                Exit Function
            End If
        Next I
    End Select
    
    For I = 1 To UBound(uSignatureG)
        If uSignatureG(I).sHash = sHash Then
            '匹配成功，返回病毒名称
            MatchSignature = uSignatureG(I).sName
            Exit Function
        End If
    Next I


End Function
