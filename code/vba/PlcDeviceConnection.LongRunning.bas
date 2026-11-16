Attribute VB_Name = "LongRunningSnippets"
Option Explicit

'DOC
Public Function ConnectAtStartup(ByVal connection As PlcDeviceConnection) As Boolean
    Call connection.Open
    Call connection.Connect

    If (Not connection.IsConnected) Then
        Debug.Print "PLC not reachable, state " & connection.State
    End If

    ConnectAtStartup = connection.IsConnected
End Function

'DOC
Public Function ReadSpeed(ByVal connection As PlcDeviceConnection) As Long
    Dim attempt As Integer

Retry:
    On Error GoTo Failed
    ReadSpeed = connection.ReadInt32("DB1.DBD 1")
    Exit Function

Failed:
    attempt = attempt + 1
    If (attempt < 2 And Err.Number = &H80131500) Then Resume Retry
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
