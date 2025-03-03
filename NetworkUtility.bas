Option Explicit

Sub EnableNetwork()
Dim networkObj As Object
Set networkObj = CreateObject("WScript.Network")
Dim strSharePath As String
Dim strDriveLetter As String

strSharePath = "\\10.1.1.28\public"
strDriveLetter = "M:" ' 你希望映射的驱动器字母

' 映射网络驱动器
On Error Resume Next ' 忽略错误，以便如果驱动器已映射则不会出错
networkObj.MapNetworkDrive strDriveLetter, strSharePath
On Error GoTo 0 ' 恢复正常的错误处理

' ... 你的代码 ...


End Sub


