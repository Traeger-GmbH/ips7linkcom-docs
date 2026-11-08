Attribute VB_Name = "WriteSnippets"
Option Explicit

'DOC
Public Sub WriteSingleValue(ByVal connection As PlcDeviceConnection)
    Call connection.WriteInt32("DB1.DBD 1", 123)
End Sub

'DOC
Public Sub WriteArray(ByVal connection As PlcDeviceConnection)
    Dim values(2) As Long
    values(0) = 123
    values(1) = 456
    values(2) = 789

    Call connection.WriteInt32Array("DB1.DBD 1", values)
End Sub

'DOC
Public Sub WriteText(ByVal connection As PlcDeviceConnection)
    Call connection.WriteString("DB111.DBB 100", "Hello World!")
End Sub
