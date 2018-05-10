VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIdentification As identification
Private pProfile As PrivateProfile
Private pTextBox As textBox

Public Sub Initialize(identification, profile, ByRef textBox As textBox)
    Set pIdentification = identification
    Set pProfile = profile
    Set pTextBox = textBox
End Sub


Public Sub LoadAndSet()
    Dim val As String
    val = pProfile.Load(pIdentification.id)
    pTextBox.Text = val
End Sub

Public Sub Store()
    Dim val As String
    val = pTextBox.Text
    Call pProfile.Store(pIdentification.id, val)
    On Error GoTo ErrXMLParts:
    Call XMLParts.StoreString(pIdentification, val)
    Exit Sub
    
ErrXMLParts:
    Debug.Print "Error number: " & Err.Number & " " & Err.Description
    Exit Sub

End Sub

