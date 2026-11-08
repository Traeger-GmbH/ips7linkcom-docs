VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'DOC
Private WithEvents m_connection As PlcDeviceConnection

Public Sub Watch(ByVal connection As PlcDeviceConnection)
    Set m_connection = connection
End Sub

'DOC
Private Sub m_connection_Connected(ByRef sender As Object, ByVal e As ComEventArgs)
    Debug.Print "Connection established"
End Sub

Private Sub m_connection_Disconnected(ByRef sender As Object, ByVal e As ComEventArgs)
    Debug.Print "Connection lost"
End Sub

'DOC
Private Sub m_connection_StateChanged( _
        ByRef sender As Object, _
        ByVal e As PlcDeviceConnectionStateChangedEventArgs)

    Debug.Print "State changed from " & e.OldState & " to " & e.NewState
End Sub
