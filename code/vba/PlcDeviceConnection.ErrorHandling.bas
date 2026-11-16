Attribute VB_Name = "ErrorHandlingSnippets"
Option Explicit

'DOC
Public Const ERR_PLC As Long = &H80131500
Public Const ERR_INVALID_ARGUMENT As Long = &H80070057
Public Const ERR_NOT_READY As Long = &H80131509
Public Const ERR_LICENSE As Long = &H80131901

Public Sub ReadWithErrorHandling(ByVal connection As PlcDeviceConnection)
    On Error GoTo Failed

    Debug.Print "Speed: " & connection.ReadInt32("DB1.DBD 1")
    Exit Sub

Failed:
    Select Case Err.Number
        Case ERR_PLC
            Debug.Print "PLC or connection: " & Err.Description
        Case ERR_INVALID_ARGUMENT
            Debug.Print "Invalid argument: " & Err.Description
        Case ERR_NOT_READY
            Debug.Print "Connection not open: " & Err.Description
        Case ERR_LICENSE
            Debug.Print "License: " & Err.Description
        Case Else
            Debug.Print "Error " & Hex$(Err.Number) & ": " & Err.Description
    End Select
End Sub

'DOC
Public Sub ReadValuesWithStatus(ByVal connection As PlcDeviceConnection)
    Dim speed As PlcInt32
    Set speed = New_PlcInt32("DB1.DBD 1")

    Dim counter As PlcInt16
    Set counter = New_PlcInt16("DB2.DBW 0")

    Dim values() As PlcValue
    values = New_PlcValues(speed, counter)

    On Error Resume Next
    Call connection.ReadValues(values)
    On Error GoTo 0

    Call PrintStatus(speed)
    Call PrintStatus(counter)
End Sub

Private Sub PrintStatus(ByVal value As PlcValue)
    Dim statusHelper As New PlcStatusHelper

    If (value.Status.Code = PlcStatusCode_NoError) Then
        Debug.Print value.Type.Address.ToString() & " = " & value.Value
    Else
        Debug.Print value.Type.Address.ToString() & ": " & value.Status.Text
        Debug.Print "Status " & value.Status.Code & ": " & statusHelper.GetText(value.Status.Code)
    End If
End Sub
