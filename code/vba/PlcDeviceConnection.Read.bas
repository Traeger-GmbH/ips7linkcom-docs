Attribute VB_Name = "ReadSnippets"
Option Explicit

'DOC
Public Sub ReadSingleValue(ByVal connection As PlcDeviceConnection)
    Dim value As Long
    value = connection.ReadInt32("DB1.DBD 1")
End Sub

'DOC
Public Sub ReadArray(ByVal connection As PlcDeviceConnection)
    Dim values() As Long
    values = connection.ReadInt32Array("DB1.DBD 1", 3)
End Sub

'DOC
Public Sub ReadText(ByVal connection As PlcDeviceConnection)
    Dim message As String
    message = connection.ReadString("DB111.DBB 100", 16)
End Sub
