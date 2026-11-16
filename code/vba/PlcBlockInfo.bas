Attribute VB_Name = "BlockInfoSnippets"
Option Explicit

'DOC
Public Sub PrintDataBlock(ByVal device As SimaticDevice)
    Dim block As PlcBlockInfo
    Set block = device.GetBlock("DB1")

    If (block.Size < 0) Then
        Debug.Print "DB1: no answer from the PLC"
    ElseIf (block.Exists) Then
        Debug.Print "DB1: " & block.Size & " bytes"
    Else
        Debug.Print "DB1 does not exist"
    End If
End Sub

'DOC
Public Sub RefreshDataBlock(ByVal block As PlcBlockInfo)
    Call block.Refresh
    Debug.Print block.Operand.ToString() & ": " & block.Size & " bytes"
End Sub
