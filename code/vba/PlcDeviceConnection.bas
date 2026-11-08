Attribute VB_Name = "PlcDeviceConnectionSnippets"
Option Explicit

'DOC
Public Sub OpenAndClose(ByVal device As SimaticDevice)
    Dim connection As PlcDeviceConnection
    Set connection = device.CreateConnection()

    Call connection.Open

    ' Read and write values of the PLC here.

    Call connection.Close
End Sub

'DOC
Public Function GetConnection(ByVal device As SimaticDevice) As PlcDeviceConnection
    Set GetConnection = device.CreateConnection(False)
End Function

'DOC
Public Sub ConfigureTimeouts(ByVal connection As PlcDeviceConnection)
    connection.ConnectTimeout = 5000
    connection.ReceiveTimeout = 5000
    connection.TransmitTimeout = 5000
End Sub

'DOC
Public Sub ConfigureBreakDetection(ByVal connection As PlcDeviceConnection)
    connection.UseBreakDetection = True
    connection.BreakDetectionTimeout = 10000
End Sub

'DOC
Public Sub PrintState(ByVal connection As PlcDeviceConnection)
    If (connection.State = PlcDeviceConnectionState_Connected) Then
        Debug.Print "Connected"
    End If

    Debug.Print connection.Status.Code & ": " & connection.Status.Text
End Sub
