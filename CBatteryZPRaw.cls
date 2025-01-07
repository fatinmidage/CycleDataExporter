'******************************************
' 类: CBatteryZPRaw
' 用途: 存储电池中检原始数据
'******************************************
Option Explicit

Private m_stepNo As Long        '工步号
Private m_batteryCode As String '电池条码
Private m_stepTime As String    '工步时间 - 使用String类型
Private m_voltage As Double     '电压
Private m_current As Double     '电流

'******************************************
' 过程: Initialize
' 用途: 初始化类的所有属性
' 参数:
'   - stepNo: 工步号
'   - batteryCode: 电池条码
'   - stepTime: 工步时间
'   - voltage: 电压
'   - current: 电流
'******************************************
Public Sub Initialize(ByVal stepNo As Long, _
                     ByVal batteryCode As String, _
                     ByVal stepTime As String, _
                     ByVal voltage As Double, _
                     ByVal current As Double)
    m_stepNo = stepNo
    m_batteryCode = batteryCode
    m_stepTime = stepTime
    m_voltage = voltage
    m_current = current
End Sub

'属性: StepNo
Public Property Get StepNo() As Long
    StepNo = m_stepNo
End Property

'属性: BatteryCode
Public Property Get BatteryCode() As String
    BatteryCode = m_batteryCode
End Property

'属性: StepTime
Public Property Get StepTime() As String
    StepTime = m_stepTime
End Property

'属性: Voltage
Public Property Get Voltage() As Double
    Voltage = m_voltage
End Property

'属性: Current
Public Property Get Current() As Double
    Current = m_current
End Property 