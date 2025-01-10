'******************************************
' 类: BatteryInfo
' 用途: 存储电池信息，包括序号和名称
'******************************************
Option Explicit

'私有变量
Private pIndex As Long        '电池序号
Private pName As String       '电池名称

'******************************************
' 属性: Index
' 用途: 获取或设置电池序号
'******************************************
Public Property Get Index() As Long
    Index = pIndex
End Property

Public Property Let Index(ByVal value As Long)
    pIndex = value
End Property

'******************************************
' 属性: Name
' 用途: 获取或设置电池名称
'******************************************
Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal value As String)
    pName = value
End Property

