Attribute VB_Name = "ModuleHashFileStream"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 取得文件节(section)的哈希值(hash)
'
'****************************************************************

Option Explicit

'API声明
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal lHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal lHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal lHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'常量定义
Private Const PROV_RSA_FULL = 1
Private Const CRYPT_NEWKEYSET = &H8
Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD5 = 3
Private Const ALGORITHM = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
Private Const HP_HASHVAL = 2
Private Const HP_HASHSIZE = 4

'****************************************************************
'
' 获得文件流hash,文件中部分内容的hash，这里是取section
' 参数：ByteStream 流(Byte数组)
' 成功返回流的hash，失败返回空
'
'****************************************************************
Public Function HashFileStream(ByteStream() As Byte) As String
    
    '初始化函数返回值
    HashFileStream = ""
    
    Dim lCtx As Long
    Dim lHash As Long
    Dim lFile As Long
    Dim lRes As Long
    Dim lLen As Long
    Dim lIdx As Long
    Dim bHash() As Byte
    Dim bBlock() As Byte
    
    Dim bStream() As Byte
    ReDim bStream(1 To UBound(ByteStream)) As Byte
    
    CopyMemory bStream(1), ByteStream(1), UBound(ByteStream)
    
    'API功能：得到CS，即密码容器
    '参数说明：
    '参数1 lCtx : 返回CSP句柄
    '参数2 : 密码容器名,为空连接缺省的CSP
    '参数3 : 为空时使用默认CSP名(微软RSA Base Provider)
    '参数4 PROV_RSA_FULLCSP : 类型
    lRes = CryptAcquireContext(lCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    
    If lRes <> 0 Then
        
        'API功能：创hash对象
        '参数说明：
        '参数1: CSP句柄
        '参数2: 选择hash算法，比如CALG_MD5等
        '参数3: HMAC 和MAC算法时有用
        '参数4: 保留，传入0即可
        '参数5:返回hash句柄
        lRes = CryptCreateHash(lCtx, ALGORITHM, 0, 0, lHash)
            
        If lRes <> 0 Then
            
            '32K 数字后加&的意思：系统默认数字是Integer型，加了&就强制把这个数转换成Long型
            Const BLOCK_SIZE As Long = 32 * 1024&
            ReDim bBlock(1 To BLOCK_SIZE) As Byte
            
            Dim lCount As Long
            Dim lBlocks As Long
            Dim lLastBlock As Long
                
            '块个数
            lBlocks = UBound(ByteStream) \ BLOCK_SIZE
            '最后一个块
            lLastBlock = UBound(ByteStream) - lBlocks * BLOCK_SIZE
        
            For lCount = 0 To lBlocks - 1
                CopyMemory bBlock(1), bStream(1 + lCount * BLOCK_SIZE), BLOCK_SIZE
                
                'API功能：hash数据
                '参数说明：
                '参数1: hash对象
                '参数2: 被hash的数据
                '参数3: 数据的长度
                '参数4: 微软的CSP这个值会被忽略
                lRes = CryptHashData(lHash, bBlock(1), BLOCK_SIZE, 0)
            Next
                
            If lLastBlock > 0 And lRes <> 0 Then
                ReDim bBlock(1 To lLastBlock) As Byte
                CopyMemory bBlock(1), bStream(1 + lCount * BLOCK_SIZE), lLastBlock
                lRes = CryptHashData(lHash, bBlock(1), lLastBlock, 0)
            End If
            
            If lRes <> 0 Then
            
                'API功能：取hash数据大小
                '参数说明：
                '参数1：hash对像
                '参数2：取hash数据大小(HP_HASHSIZE)
                '参数3：返回hash数据长度
                lRes = CryptGetHashParam(lHash, HP_HASHSIZE, lLen, 4, 0)
                
                If lRes <> 0 Then
                    ReDim bHash(0 To lLen - 1)
                    
                    'API功能：取得hash数据
                    '参数说明：
                    '参数1：hash对像
                    '参数2：取hash数据(值)(HP_HASHVAL)
                    '参数3：hash数组
                    lRes = CryptGetHashParam(lHash, HP_HASHVAL, bHash(0), lLen, 0)
                    If lRes <> 0 Then
                        For lIdx = 0 To UBound(bHash)
                            '构造hash值，2位一组，不足补0
                            HashFileStream = HashFileStream & Right("0" & Hex(bHash(lIdx)), 2)
                        Next
                    End If
                End If
            End If
    
            CryptDestroyHash lHash
        End If
    End If
    CryptReleaseContext lCtx, 0

End Function

'****************************************************************
'
' 获得文件hash值
' 参数：文件名
' 成功返回文件的hash，失败返回空
' 代码与HashFileStream相似，注释可参见HashFileStream
'
'****************************************************************
Function HashFile(ByVal sFilename As String) As String
    
    HashFile = ""
    
    Dim lCtx As Long
    Dim lHash As Long
    Dim lFile As Long
    Dim lRes As Long
    Dim lLen As Long
    Dim lIdx As Long
    Dim bHash() As Byte
    
    If Len(Dir(sFilename)) = 0 Then
        Exit Function
    End If
    lRes = CryptAcquireContext(lCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes <> 0 Then
        lRes = CryptCreateHash(lCtx, ALGORITHM, 0, 0, lHash)
        If lRes <> 0 Then
            lFile = FreeFile
            Open sFilename For Binary As #lFile
                Const BLOCK_SIZE As Long = 32 * 1024&
                ReDim bBlock(1 To BLOCK_SIZE) As Byte
                Dim lCount As Long
                Dim lBlocks As Long
                Dim lLastBlock As Long
                lBlocks = LOF(lFile) \ BLOCK_SIZE
                lLastBlock = LOF(lFile) - lBlocks * BLOCK_SIZE
                For lCount = 1 To lBlocks
                    Get lFile, , bBlock
                    lRes = CryptHashData(lHash, bBlock(1), BLOCK_SIZE, 0)
                Next
                If lLastBlock > 0 And lRes <> 0 Then
                    ReDim bBlock(1 To lLastBlock) As Byte
                    Get lFile, , bBlock
                    lRes = CryptHashData(lHash, bBlock(1), lLastBlock, 0)
                End If
            Close lFile
            If lRes <> 0 Then
                lRes = CryptGetHashParam(lHash, HP_HASHSIZE, lLen, 4, 0)
                If lRes <> 0 Then
                    ReDim bHash(0 To lLen - 1)
                    lRes = CryptGetHashParam(lHash, HP_HASHVAL, bHash(0), lLen, 0)
                    If lRes <> 0 Then
                        For lIdx = 0 To UBound(bHash)
                            HashFile = HashFile & Right("0" & Hex(bHash(lIdx)), 2)
                        Next
                    End If
                End If
            End If
            CryptDestroyHash lHash
        End If
    End If
    CryptReleaseContext lCtx, 0
End Function
