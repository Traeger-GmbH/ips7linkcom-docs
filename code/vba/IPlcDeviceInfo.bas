Attribute VB_Name = "DeviceInfoSnippets"
Option Explicit

'DOC
Public Sub PrintDeviceInfo(ByVal device As SimaticDevice)
    Dim info As IPlcDeviceInfo
    Set info = device.GetInfo()

    If (info.HasName) Then Debug.Print "Name: " & info.Name
    If (info.HasLocation) Then Debug.Print "Location: " & info.Location
    If (info.HasPlantId) Then Debug.Print "Plant ID: " & info.PlantId
    If (info.HasModuleName) Then Debug.Print "Module: " & info.ModuleName
    If (info.HasModuleType) Then Debug.Print "Module type: " & info.ModuleType
    If (info.HasModuleSerial) Then Debug.Print "Serial number: " & info.ModuleSerial
    If (info.HasMemoryCardSerial) Then Debug.Print "Memory card: " & info.MemoryCardSerial
    If (info.HasCopyright) Then Debug.Print "Copyright: " & info.Copyright
    If (info.HasTime) Then Debug.Print "Time: " & info.Time
End Sub
