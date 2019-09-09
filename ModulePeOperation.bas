Attribute VB_Name = "ModulePeOperation"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' PE文件操作,扫描文件是否是病毒
'
'****************************************************************
Option Explicit

Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSize As LARGE_INTEGER) As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

'****************************************************************
'
'PE文件结构
'
'+---------------------+
'|   1 DOS头
'+---------------------+
'|   2 文件头
'+---------------------+
'|   3 可选文件头
'|--------------------
'|   |-----------------+
'|   | (数据目录)Data Directories
'|   | | (RVA,Size)
'+---------------------+
'|   4(节头)Section-Headers
'+---------------------+
'|   5(各节数据)Section   Data
'+---------------------+
'****************************************************************

'****************************************************************
'(1)
'DOS头
'除了最后的lfanew和e_magic基本都可以随便填
'这个部分，我们关心的为 lfanew,
'该地址指向了PE头结构在文件中的偏移.如我们的这个文件该值为0x40,而0x40开始就是PE头的开始,标志为"PE"两个字符.
Private Type IMAGE_DOS_HEADER
    Magic As Integer
    cblp As Integer
    cp As Integer
    crlc As Integer
    cparhdr As Integer
    minalloc As Integer
    maxalloc As Integer
    ss As Integer
    sp As Integer
    csum As Integer
    ip As Integer
    cs As Integer
    lfarlc As Integer
    ovno As Integer
    res(3) As Integer
    oemid As Integer
    oeminfo As Integer
    res2(9) As Integer
    lfanew As Long
End Type

'(2)
'文件头
Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOtionalHeader As Integer
    Characteristics As Integer
End Type

'(3)
'可选文件头
Private Type IMAGE_DATA_DIRECTORY
    '每个地址都是Virtual,就是程序加载后该地址在内存中相对于ImageBase的偏移,带virtual的地址是这样
    DataRVA As Long
    '在文件中的偏移为：Addr - 对应Section的VirtualAddress + 对应Section的文件内起始偏移
    DataSize As Long
End Type
Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkVer As Byte
    MinorLinkVer As Byte
    CodeSize As Long
    InitDataSize As Long
    unInitDataSize As Long
    '入口地址,佷重要(代码节的开始处)
    EntryPoint As Long
    '代码段起始地址(RVA)
    CodeBase As Long
    '数据段起始地址(RVA)
    DataBase As Long
    'NT/2k/XP基本都是0x400000
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOSVer As Integer
    MinorOSVer As Integer
    MajorImageVer As Integer
    MinorImageVer As Integer
    MajorSSVer As Integer
    MinorSSVer As Integer
    Win32Ver As Long
    ImageSize As Long
    HeaderSize As Long
    Checksum As Long
    '这个别乱填,要不然可能又是非法Win32程序,MajorSubsystemVersion也是
    Subsystem As Integer
    DLLChars As Integer
    StackRes As Long
    StackCommit As Long
    HeapReserve As Long
    HeapCommit  As Long
    LoaderFlags As Long
    RVAsAndSizes As Long
    DataEntries(15) As IMAGE_DATA_DIRECTORY
End Type

'****************************************************************
'15个数据目录：
'1   Export table address and size(输出表)
'2   Import table address and size(输入表)
'3   Resource table address and size
'4   Exception   table   address   and   size
'5   Certificate   table   address   and   size
'6   Base   relocation   table   address   and   size
'7   Debugging   information   starting   address   and   size
'8   Architecture-specific   data   address   and   size
'9   Global   pointer   register   relative   virtual   address
'10  Thread   local   storage   (TLS)   table   address   and   size
'11  Load   configuration   table   address   and   size
'12  Bound   import   table   address   and   size
'13  Import   address   table   address   and   size
'14  Delay   import   descriptor   address   and   size
'15  Reserved
'16  Reserved
'****************************************************************

'(4)
'节头
Private Type IMAGE_SECTION_HEADER
    '节名(8 byte)
    SectionName(7) As Byte
    VirtualSize As Long
    '节的RVA
    VirtualAddress As Long
    '节大小 Raw Data
    SizeOfRawData As Long
    '节的文件偏移量
    PointerToRawData As Long
    PointerToRelocations As Long
    PLineNums As Long
    RelocCount As Integer
    LineCount As Integer
    '节的属性,如可读可写
    Characteristics As Long
End Type

'(5)
'节数据

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const OPEN_EXISTING = 3

'****************************************************************
'
' 扫描文件Section，以hash(节大小+节hash)进行特征码匹配检测病毒
' 特别说明：扫描文件入口地址所在节
' 参数：hFile 文件句柄,传入前检查合法性
' 返回值：检测到返回病毒名，否则返回空
'
'****************************************************************
Private Function ScanSections(hFile As Long) As String
    
    '初始化返回
    ScanSections = ""
    
    'DOS头
    Dim uDos          As IMAGE_DOS_HEADER
    '文件头
    Dim uFile         As IMAGE_FILE_HEADER
    '可选头
    Dim uOptional     As IMAGE_OPTIONAL_HEADER
    '节头
    Dim uSections()   As IMAGE_SECTION_HEADER
    
    Dim lBytesRead As Long
    Dim I As Long
    Dim bTempMemory() As Byte
    

    '把指针移动到文件头
    SetFilePointer hFile, 0, 0, 0
    
    '读取DOS头
    ReadFile hFile, uDos, Len(uDos), lBytesRead, ByVal 0&
    
    '根据DOS头的lfanew，指向PE头+4的地方(即文件头开始地址)
    SetFilePointer hFile, ByVal uDos.lfanew + 4, 0, 0
    
    '读取文件头
    ReadFile hFile, uFile, Len(uFile), lBytesRead, ByVal 0&
        
    '读取可选头
    ReadFile hFile, uOptional, Len(uOptional), lBytesRead, ByVal 0&
    
    '判断节数是否正确
    If uFile.NumberOfSections < 1 Or uFile.NumberOfSections > 99 Then
        Exit Function
    End If
    
    '重写义节数(uFile.NumberOfSections是节个数)
    ReDim uSections(uFile.NumberOfSections - 1) As IMAGE_SECTION_HEADER
    
    '读取节头
    ReadFile hFile, uSections(0), Len(uSections(0)) * uFile.NumberOfSections, lBytesRead, ByVal 0&
    
    '扫描每个节
    For I = 0 To UBound(uSections)
    
        '防止CPU占用过高及程序无响应
        DoEvents
            
        'SizeOfRawData有为0的时候,如果为0就要出错
        If uSections(I).SizeOfRawData > 0 Then
            
            '最大段不超过10MB
            If (uSections(I).SizeOfRawData / 1024 / 1024) <= 10 Then
    
                ReDim bTempMemory(1 To uSections(I).SizeOfRawData) As Byte
                    
                '设置指针到节的起始处
                SetFilePointer hFile, ByVal uSections(I).PointerToRawData, 0, 0
                    
                '读取整节数据,读取后TemoMemory里存放的是要hash的数据，完整的section
                ReadFile hFile, bTempMemory(1), uSections(I).SizeOfRawData, lBytesRead, ByVal 0&
                    
                Dim sTargetStr As String
                    
                '构造特征码：节大小+节的hash值
                sTargetStr = uSections(I).SizeOfRawData & ":" & LCase(HashFileStream(bTempMemory))
                    
                Dim bTempStrByte() As Byte
                bTempStrByte = sTargetStr
                    
                '用哈希算法加密特征码(hash)
                sTargetStr = HashFileStream(bTempStrByte)
                    
                '特征码匹配检测
                Dim sMatchResult As String
                sMatchResult = MatchSignature(sTargetStr)
                    
                '如果不为空表示检测到是病毒，返回
                If sMatchResult <> "" Then
                    ScanSections = Trim(sMatchResult)
                    Exit Function
                End If

            End If
        End If
    Next I
    
End Function

'****************************************************************
'
' 扫描PE文件
' 参数：sFile 文件(带全路径)
' 返回值：扫描到病毒返回病毒名，没扫到返回空
'
'****************************************************************
Public Function ScanFile(sFile As String) As String
    
    '初始化返回值
    ScanFile = ""
    
    
    '判断文件大小，如果文件太大，则不可能是病毒，直接放行
    Dim lFileSize As Long
    Dim uSize As LARGE_INTEGER
    
    '打开文件
    lFileSize = CreateFile(sFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    '获取文件大小
    GetFileSizeEx lFileSize, uSize
    '关闭文件
    CloseHandle lFileSize
    If uSize.lowpart / 1024 / 1024 > 15 Then
        Exit Function
    End If
        
    '打开文件
    Dim lFileHwnd As Long
    lFileHwnd = CreateFile(sFile, ByVal &H80000000, 0, ByVal 0&, 3, 0, ByVal 0)
    
    If lFileHwnd Then
            
        '打开文件，检查是否是PE文件
        Dim bBuffer(12) As Byte
        Dim lBytesRead As Long
        Dim uDosHeader As IMAGE_DOS_HEADER
    
        ReadFile lFileHwnd, uDosHeader, ByVal Len(uDosHeader), lBytesRead, ByVal 0&
        CopyMemory bBuffer(0), uDosHeader.Magic, 2
        
        '第一个检查点
        If (Chr(bBuffer(0)) & Chr(bBuffer(1)) = "MZ") Then
                    
            '把指针设置到PE头结构偏移处
            SetFilePointer lFileHwnd, uDosHeader.lfanew, 0, 0
            
            '读取4byte
            ReadFile lFileHwnd, bBuffer(0), 4, lBytesRead, ByVal 0&
            '检查是否是PE(回车)(回车)
            If (Chr(bBuffer(0)) = "P") And (Chr(bBuffer(1)) = "E") And (bBuffer(2) = 0) And (bBuffer(3) = 0) Then
                
                '以Section方式扫描文件,成功返回病毒名，失败返回空
                Dim sScanSectionResult As String
                sScanSectionResult = ScanSections(lFileHwnd)
                
                '如果返回不为空，表示检测到是病毒，则关闭文件，返回
                If sScanSectionResult <> "" Then
                    ScanFile = sScanSectionResult
                    CloseHandle lFileHwnd
                    Exit Function
                End If
            End If
        End If
    End If
    CloseHandle lFileHwnd
    
End Function

