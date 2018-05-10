VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DTPickerConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIdentification As identification
Private pProfile As PrivateProfile
Private pDateBox As Control


Public Sub Initialize(identification, profile, ByRef dateBox As Control)
    Set pIdentification = identification
    Set pProfile = profile
    Set pDateBox = dateBox
End Sub


Public Sub LoadAndSet()
    Dim val As String
    val = pProfile.LoadDate(pIdentification.id)
    pDateBox.Value = val
End Sub

Public Sub Store()
    Dim val As Date
    val = pDateBox.Value
    Call pProfile.StoreDate(pIdentification.id, val)
    
    On Error GoTo ErrXMLParts:
    Call XMLParts.StoreDate(pIdentification, val)
    
    Exit Sub
    
ErrXMLParts:
    Debug.Print "Error number: " & Err.Number & " " & Err.Description
    Exit Sub
    
End Sub


