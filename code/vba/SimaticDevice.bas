Attribute VB_Name = "SimaticDeviceSnippets"
Option Explicit

'DOC
Public Function CreateDevice() As SimaticDevice
    Set CreateDevice = New_SimaticDevice( _
            New_IPDeviceEndPoint("192.168.0.80"), SimaticDeviceType_S71500)
End Function

'DOC
Public Function CreateDeviceWithRackAndSlot() As SimaticDevice
    Set CreateDeviceWithRackAndSlot = New_SimaticDevice( _
            New_IPDeviceEndPoint("192.168.0.80", 0, 2), SimaticDeviceType_S7300_400)
End Function

'DOC
Public Function CreateDeviceWithChannelType() As SimaticDevice
    Dim device As SimaticDevice
    Set device = New_SimaticDevice( _
            New_IPDeviceEndPoint("192.168.0.80"), SimaticDeviceType_S71200)

    device.ChannelType = SimaticChannelType_OperationPanel
    Set CreateDeviceWithChannelType = device
End Function

'DOC
Public Function CreateDeviceWithoutHelper() As SimaticDevice
    Dim endPointFactory As New IPDeviceEndPointFactory
    Dim deviceFactory As New SimaticDeviceFactory

    Set CreateDeviceWithoutHelper = deviceFactory.Create( _
            endPointFactory.Create("192.168.0.80"), SimaticDeviceType_S71500)
End Function
