Attribute VB_Name = "GettingStarted"
Option Explicit

'DOC
Public Sub ReadTemperature()
    Dim device As SimaticDevice
    Set device = New_SimaticDevice( _
            New_IPDeviceEndPoint("192.168.0.80"), SimaticDeviceType_S7300_400)

    Dim connection As PlcDeviceConnection
    Set connection = device.CreateConnection()

    Call connection.Open

    Dim temperature As Single
    temperature = connection.ReadReal("DB10.DBD 20")
    Debug.Print "Current Temperature is " & temperature & " C"

    Call connection.Close
End Sub
