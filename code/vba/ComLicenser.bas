Attribute VB_Name = "LicenseSnippets"
Option Explicit

'DOC
Public Sub SetLicenseKey()
    Dim licenser As New ComLicenser
    licenser.LicenseKey = "<insert your license code here>"
End Sub

'DOC
Public Sub LoadLicenseKeyFromFile()
    Dim licenser As New ComLicenser
    Call licenser.LoadLicenseKey("C:\Licenses\license.xml")
End Sub

'DOC
Public Sub PrintLicense()
    Dim licenser As New ComLicenser

    If (licenser.IsEvaluation) Then
        Debug.Print "Evaluation"
    Else
        Debug.Print "Licensed until " & licenser.ExpiryDate
        Debug.Print "Support until " & licenser.SupportExpiryDate
    End If
End Sub
