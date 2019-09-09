Attribute VB_Name = "ModuleGetFileIcon"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 获取文件图标
'
'****************************************************************

Option Explicit

'API声明
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'常量定义
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_bSmallIcon = &H1

'自定义类型
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSID
    id(16) As Byte
End Type

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 255
    szTypeName As String * 80
End Type

'****************************************************************
'
' ICON 转 Picture
' 参数： hIcon 图标句柄
'
'****************************************************************
Public Function IconToPicture(hIcon As Long) As IPictureDisp

    Dim cls_id As CLSID
    Dim hRes As Long
    Dim new_icon As TypeIcon
    Dim lpUnk As IUnknown

    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    If hRes = 0 Then
        Set IconToPicture = lpUnk
    End If
    
End Function

'****************************************************************
'
' 获得文件ICON
' 参数：sFilename 文件名(全路径)
'       bSmallIcon boolean型，传入true为16*16小图标，传入false
'       为32*32大图标
'
'****************************************************************
Public Function GetFileIcon(sFilename, ByVal bSmallIcon As Boolean) As IPictureDisp


    Dim hIcon As Long
    Dim item_num As Long
    Dim icon_pic As IPictureDisp
    Dim sh_info As SHFILEINFO
    
    If bSmallIcon = True Then
        SHGetFileInfo sFilename, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_bSmallIcon
    Else
        SHGetFileInfo sFilename, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_LARGEICON
    End If
    
    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetFileIcon = icon_pic
    
End Function
