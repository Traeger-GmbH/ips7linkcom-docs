Attribute VB_Name = "PlcValueSnippets"
Option Explicit

'DOC
Public Sub ReadValueObjects(ByVal connection As PlcDeviceConnection)
    Dim speed As PlcInt32
    Set speed = New_PlcInt32("DB1.DBD 1")

    Dim coordinates As PlcInt32Array
    Set coordinates = New_PlcInt32Array("DB1.DBD 5", 3)

    Dim values As New PlcValueCollection
    Call values.Add(speed)
    Call values.Add(coordinates)

    Call connection.ReadValuesUsingCollection(values)

    Debug.Print "Speed: " & speed.Value
End Sub

'DOC
Public Sub WriteValueObjects(ByVal connection As PlcDeviceConnection)
    Dim speed As PlcInt32
    Set speed = New_PlcInt32("DB1.DBD 1", 1200)

    Dim jobIsActive As PlcBoolean
    Set jobIsActive = New_PlcBoolean("DB1.DBX 17.0", True)

    Dim values As New PlcValueCollection
    Call values.Add(speed)
    Call values.Add(jobIsActive)

    Call connection.WriteValuesUsingCollection(values)
End Sub

'DOC
Public Sub CheckStatus(ByVal speed As PlcInt32)
    If (speed.Status.Code = PlcStatusCode_NoError) Then
        Debug.Print "Speed: " & speed.Value
    Else
        Debug.Print "Speed: " & speed.Status.Text
    End If
End Sub
