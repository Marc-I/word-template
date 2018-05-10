VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuildingBlockConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Connector for Building Blocks
'
' Available values are taken from building blocks
' that the building block content control displays

' https://msdn.microsoft.com/en-us/vba/word-vba/articles/working-with-building-blocks


Private pIdentification As ContentControlIdentification
Private pProfile As PrivateProfile
Private pBuildingBlock As buildingBlock
Private pContentControl As contentControl
Private pComboBox As combobox
' Keys and Values are the building block names (both key and value are the same)
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
    
        Call LoadValuesFromBuildingBlock
        Call ConfigureComboboxFromValues
        
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

Public Function GetContentControl() As contentControl
    GetContentControl = Application.ActiveDocument.SelectContentControlsByTag(pIdentification.id).Item(1)
End Function

Public Property Get ContentControlExists() As Boolean
    ContentControlExists = Application.ActiveDocument.SelectContentControlsByTag(pIdentification.id).Count > 0
End Property

Private Function GetTemplate() As template
    ' There seem to be other ways to get a/the template:
    ' Set template = ActiveDocument.AttachedTemplate
    ' Set template = Templates("Mastervorlage.dotm")
    
    ' ' https://msdn.microsoft.com/en-us/vba/word-vba/articles/application-activedocument-property-word
    Set GetTemplate = ActiveDocument.AttachedTemplate
End Function

Private Sub LoadValuesFromBuildingBlock()
    Set pContentControl = Application.ActiveDocument.SelectContentControlsByTag(pIdentification.id).Item(1)
    If pContentControl.Type = wdContentControlBuildingBlockGallery Then
        ' Get all BuildingBlocks with given Type
        Dim template As template
        Dim buildingBlock As buildingBlock
        
        Set template = GetTemplate()
        
        For intCount = 1 To template.BuildingBlockEntries.Count
            Set buildingBlock = template.BuildingBlockEntries(intCount)
            If buildingBlock.category.name = pContentControl.BuildingBlockCategory Then
                ' Debug.Print buildingBlock.Name
                ' Debug.Print buildingBlock.id
                Call pKeysAndValues.Add(buildingBlock.name, buildingBlock.name)
            End If
        Next
    End If
End Sub

Public Sub ConfigureComboboxFromValues()
    ' https://msdn.microsoft.com/en-us/vba/access-vba/articles/combobox-removeitem-method-access
    pComboBox.Clear
    For Each keyAndValue In pKeysAndValues.KeyValuePairs
        ' https://msdn.microsoft.com/en-us/vba/access-vba/articles/combobox-additem-method-access
        Call pComboBox.AddItem(keyAndValue.Key)
    Next
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
        
        ' Debug.Print pContentControl.BuildingBlockType
        ' Debug.Print pContentControl.PlaceholderText
        ' Debug.Print pContentControl.Range.Delete
        pContentControl.Range.Delete
        
        Dim template As template
        Set template = GetTemplate()

        ' http://www.vbaexpress.com/forum/showthread.php?59163-BuildingBlock-ContentControl-Select-Item-to-Display-with-VBA
        ' http://www.msofficeforums.com/word-vba/18622-how-make-building-block-content-control-determine.html
        
        Dim bb As buildingBlock
        Set bb = GetBuildingBlock(pContentControl.BuildingBlockCategory, val)
        If bb Is Nothing Then
           Return
        End If
        
        '
        Dim contentControl As contentControl
        For Each contentControl In Application.ActiveDocument.SelectContentControlsByTag(pIdentification.id)
            bb.Insert pContentControl.Range
        Next
        
    Else
        Call pProfile.Store(pIdentification.id, pComboBox.Text)
    End If
    Exit Sub
End Sub

Private Function GetBuildingBlock(category As String, name As String) As buildingBlock
    Dim template As template
    Set template = GetTemplate()
    For intCount = 1 To template.BuildingBlockEntries.Count
        Dim buildingBlock As buildingBlock
        Set buildingBlock = template.BuildingBlockEntries(intCount)
        If buildingBlock.category.name = category Then
            If buildingBlock.name = name Then
                Set GetBuildingBlock = buildingBlock
                Exit Function
            End If
        End If
    Next
    Set GetBuildingBlock = Nothing
End Function

