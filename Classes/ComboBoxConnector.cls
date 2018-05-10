VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pIdentification As identification
Private pProfile As PrivateProfile
Private pComboBox As combobox
Private pKeysAndValues As Dictionary

Public Sub Initialize(identification, profile, ByRef c As combobox)
    Set pIdentification = identification
    Set pProfile = profile
    Set pComboBox = c
    Set pKeysAndValues = New Dictionary
End Sub


Public Sub LoadAndSetComboBox()
    Dim val As String
    If ContentControlExists() Then
        
        val = pProfile.Load(pIdentification.id)
    
        Call LoadValuesFromContentControl
        
        ' Adding Elements to a Combobox control: https://msdn.microsoft.com/en-us/library/aa240832(v=vs.60).aspx
        For Each keyAndValue In pKeysAndValues.KeyValuePairs
            Call pComboBox.AddItem(keyAndValue.Value)
        Next
        
        selectedIndex = pKeysAndValues.IndexOfKey(val)
        If selectedIndex <> -1 Then
            pComboBox.ListIndex = selectedIndex
            Exit Sub
        End If
        selectedIndex = pKeysAndValues.IndexOfValue(val)
        If selectedIndex <> -1 Then
            pComboBox.ListIndex = selectedIndex
            Exit Sub
        End If
        pComboBox.Text = val
    Else
        val = pProfile.Load(pIdentification.id)
        pComboBox.Text = val
    End If
End Sub

Public Property Get ContentControlExists() As Boolean
    ContentControlExists = Application.ActiveDocument.SelectContentControlsByTag(pIdentification.id).Count > 0
End Property

Private Sub LoadValuesFromContentControl()
    Dim contentControl As contentControl
    Set contentControl = Application.ActiveDocument.SelectContentControlsByTag(pIdentification.id).Item(1)
    For Each entry In contentControl.DropdownListEntries
        If Len(Trim(entry.Value)) <> 0 Then
            Call pKeysAndValues.Add(entry.Value, entry.Text)
            ' Debug.Print entry.Value, entry.Text
        End If
    Next entry
End Sub

Public Property Get Value() As String
    Value = pKeysAndValues.Key(pComboBox.ListIndex)
End Property

Public Sub Store()
    If ContentControlExists() Then
        Dim val As String
        ' Differenz between property Value and Text of a Combobox control: https://msdn.microsoft.com/de-de/VBA/Access-VBA/articles/combobox-value-property-access
        val = pComboBox.Value()
        Call pProfile.Store(pIdentification.id, val)
        On Error GoTo ErrXMLParts:
        Call XMLParts.StoreString(pIdentification, val)
    Else
        Call pProfile.Store(pIdentification.id, pComboBox.Text)
    End If
    Exit Sub
    
ErrXMLParts:
    Debug.Print "Error number: " & Err.Number & " " & Err.Description
    Exit Sub

End Sub
