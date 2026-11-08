Attribute VB_Name = "PlcAddressSnippets"
Option Explicit

'DOC
Public Sub ParseAddresses()
    Dim address As PlcAddress
    Set address = PlcAddresses.Parse("DB100.DBX 1.0")

    Dim inputAddress As PlcAddress
    If (PlcAddresses.TryParse("E 1.0", inputAddress)) Then
        Debug.Print inputAddress.ToString(PlcOperandStandard_IEC)
    End If
End Sub

'DOC
Public Sub CreateAddresses()
    Dim dataBlockAddress As PlcAddress
    Set dataBlockAddress = New_PlcAddress(PlcOperands.DataBlock(100), PlcRawType_Bit, 1, 0)

    Dim inputAddress As PlcAddress
    Set inputAddress = New_PlcAddress(PlcOperands.Input, PlcRawType_Byte, 1)
End Sub
