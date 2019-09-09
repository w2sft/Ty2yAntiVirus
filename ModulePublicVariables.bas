Attribute VB_Name = "ModulePublicVariables"
'****************************************************************
'
' Ty2y杀毒软件
' http://www.ty2y.com/
'
' 共公变量
'
'****************************************************************
Option Explicit

'主动防御接收消息窗体子类化旧地址
Public lSubClassOldAddress As Long

'是否可以卸载窗体
Public bEnableUnloadForm As Boolean

'是否进行自定义扫描标识
Public bDoCustmerScan As Boolean

'皮肤
Public sSkin As String
