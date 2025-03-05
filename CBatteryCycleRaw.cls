'******************************************
' 类: CBatteryCycleRaw
' 用途: 存储电池循环测试数据
'******************************************
Option Explicit

Private pStepNo As Long        '工步号
Private pCapacity As Double    '容量
Private pEnergy As Double      '能量
Private pBatteryCode As String '电池条码
Private pCycleNumber As Long   '循环圈数

'******************************************
' 过程: Initialize
' 用途: 初始化类的所有属性
' 参数:
'   - stepNo: 工步号
'   - capacity: 容量
'   - energy: 能量
'   - batteryCode: 电池条码（可选）
'   - cycleNumber: 循环圈数（可选）
'******************************************
Public Sub Initialize(ByVal stepNo As Long, _
                     ByVal capacity As Double, _
                     ByVal energy As Double, _
                     Optional ByVal batteryCode As String = "", _
                     Optional ByVal cycleNumber As Long = 0)
    pStepNo = stepNo
    pCapacity = capacity
    pEnergy = energy
    pBatteryCode = batteryCode
    pCycleNumber = cycleNumber
End Sub

'******************************************
' 属性: StepNo
' 用途: 获取或设置工步号
'******************************************
Public Property Get StepNo() As Long
    StepNo = pStepNo
End Property
Public Property Let StepNo(value As Long)
    pStepNo = value
End Property

'******************************************
' 属性: Capacity
' 用途: 获取或设置容量
'******************************************
Public Property Get Capacity() As Double
    Capacity = pCapacity
End Property
Public Property Let Capacity(value As Double)
    pCapacity = value
End Property

'******************************************
' 属性: Energy
' 用途: 获取或设置能量
'******************************************
Public Property Get Energy() As Double
    Energy = pEnergy
End Property
Public Property Let Energy(value As Double)
    pEnergy = value
End Property

'******************************************
' 属性: BatteryCode
' 用途: 获取或设置电池条码
'******************************************
Public Property Get BatteryCode() As String
    BatteryCode = pBatteryCode
End Property
Public Property Let BatteryCode(value As String)
    pBatteryCode = value
End Property

'******************************************
' 属性: CycleNumber
' 用途: 获取或设置循环圈数
'******************************************
Public Property Get CycleNumber() As Long
    CycleNumber = pCycleNumber
End Property
Public Property Let CycleNumber(value As Long)
    pCycleNumber = value
End Property 