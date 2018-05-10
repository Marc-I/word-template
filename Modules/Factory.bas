Attribute VB_Name = "Factory"

Public Function CreateIdentification(id As String, elementName As String, namespace As String) As identification
    Dim ident As New identification
    Call ident.Initialize(id, elementName, namespace)
    Set CreateIdentification = ident
End Function

Public Function CreateContentControlIdentification(id As String, tag As String) As ContentControlIdentification
    Dim ident As New ContentControlIdentification
    Call ident.Initialize(id, tag)
    Set CreateContentControlIdentification = ident
End Function


Public Function CreateComboBoxConnector(id As identification, profile As PrivateProfile, ByRef combobox As combobox) As ComboBoxConnector
    Dim connector As New ComboBoxConnector
    Call connector.Initialize(id, profile, combobox)
    Set CreateComboBoxConnector = connector
End Function

Public Function CreateTextBoxConnector(id As identification, profile As PrivateProfile, ByRef textBox As textBox) As TextBoxConnector
    Dim connector As New TextBoxConnector
    Call connector.Initialize(id, profile, textBox)
    Set CreateTextBoxConnector = connector
End Function

Public Function CreateDTPickerConnector(id As identification, profile As PrivateProfile, ByRef dtPicker As Control) As DTPickerConnector
    Dim connector As New DTPickerConnector
    Call connector.Initialize(id, profile, dtPicker)
    Set CreateDTPickerConnector = connector
End Function

Public Function CreateBuildingBlockConnector(id As ContentControlIdentification, profile As PrivateProfile, ByRef combobox As combobox) As BuildingBlockConnector
    Dim connector As New BuildingBlockConnector
    Call connector.Initialize(id, profile, combobox)
    Set CreateBuildingBlockConnector = connector
End Function
