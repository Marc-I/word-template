Attribute VB_Name = "ScrollOffice"
Public Function StandardDictionary() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    Call dic.Add("title", "$scroll.title")
    Set StandardDictionary = dic
End Function

Public Function PagePropertiesDictionary() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    Call dic.Add("title", "$scroll.title")
    Call dic.Add("author", "$scroll.pageproperty.(Autor)")
    Call dic.Add("issuingOffice", "$scroll.pageproperty.(Ausgabestelle)")
    Call dic.Add("scope", "$scroll.pageproperty.(Geltungsbereich)")
    Call dic.Add("classification", "$scroll.pageproperty.(Klassifizierung)")
    Call dic.Add("version", "$scroll.pageproperty.(Version)")
    Call dic.Add("issuingDate", "$scroll.pageproperty.(Ausgabedatum)")
    Call dic.Add("distribution", "$scroll.pageproperty.(Verteiler)")
    Set PagePropertiesDictionary = dic
End Function

Public Function ConfluenceDictionary() As Dictionary
    Dim dic As Dictionary
    Set dic = New Dictionary
    Call dic.Add("title", "$scroll.title")
    Call dic.Add("author", "$scroll.modifier.fullName")
    Call dic.Add("issuingOffice", "$scroll.space.name")
    Call dic.Add("scope", "$scroll.space.name")
    Call dic.Add("classification", "Intern")
    Call dic.Add("version", "$scroll.version")
    Call dic.Add("issuingDate", "$scroll.modificationdate")
    Call dic.Add("distribution", "-")
    Set ConfluenceDictionary = dic
End Function


Public Sub Replace(ByRef cc As contentControl, ByRef dic As Dictionary)
    Dim tV As Variant
    tV = cc.tag
    
    Dim t As String
    t = CStr(tV)
    Debug.Print (t)
    
    Dim v As String
    If dic.Exists(t) Then
       
       v = dic.Item(t)
    
       Dim r As Range
       Set r = cc.Range
           
       cc.Delete
       r.Delete
       r.InsertAfter (v)
    End If
End Sub
    
    
Sub ReplaceContentControls(ByRef dic As Dictionary)
    Dim doc As Document
    Set doc = Application.ActiveDocument
    
    ' https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm
    Dim rngStory As Word.Range
    Dim lngJunk As Long
    'Fix the skipped blank Header/Footer problem as provided by Peter Hewett
    lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
    'Iterate through all story types in the current document
    For Each rngStory In ActiveDocument.StoryRanges
        'Iterate through all linked stories
        Do
            Dim cc As contentControl
            For Each cc In rngStory.ContentControls
                Call Replace(cc, dic)
            Next
            'Get next linked story (if any)
            Set rngStory = rngStory.NextStoryRange
        Loop Until rngStory Is Nothing
    Next
End Sub
    
Sub Replace3EndWithScrollContent()
    '
    Call Selection.GoTo(wdGoToPage, wdGoToAbsolute, 3)
    Call Selection.EndKey(wdStory, wdExtend)
    Dim r As Range
    Set r = Selection.Range
    r.Delete
    r.InsertAfter ("$scroll.content")

End Sub

Sub ConvertToPageProperties()
    Dim dic As Dictionary
    Set dic = PagePropertiesDictionary()
    
    Call ReplaceContentControls(dic)
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls(dic)
    Call ReplaceContentControls(dic)
    
    Call Replace3EndWithScrollContent
    
    Call CreateScrollOfficeStyles
  
End Sub

Sub ConvertToStandard()
    Dim dic As Dictionary
    Set dic = StandardDictionary()
    
    Call ReplaceContentControls(dic)
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls(dic)
    Call ReplaceContentControls(dic)

    Call Replace3EndWithScrollContent
    
    Call CreateScrollOfficeStyles
End Sub

Sub ConvertToConfluence()
Dim dic As Dictionary
    Set dic = ConfluenceDictionary()
    
    Call ReplaceContentControls(dic)
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls(dic)
    Call ReplaceContentControls(dic)

    Call Replace3EndWithScrollContent
    
    Call CreateScrollOfficeStyles

End Sub

Sub CreateScrollOfficeStyles()
    Dim doc As Document
    Set doc = Application.ActiveDocument
    
    Call CreateOrEditStyle("Scroll List Bullet", "Aufz�hlungszeichen")
    Call CreateOrEditStyle("Scroll List Bullet 1", "Aufz�hlungszeichen")
    Call CreateOrEditStyle("Scroll List Bullet 2", "Aufz�hlungszeichen 2")
    Call CreateOrEditStyle("Scroll List Bullet 3", "Aufz�hlungszeichen 3")
    
    Call CreateOrEditStyle("Scroll List Number", "Listennummer")
    Call CreateOrEditStyle("Scroll List Number 1", "Listennummer")
    Call CreateOrEditStyle("Scroll List Number 2", "Listennummer 2")
    Call CreateOrEditStyle("Scroll List Number 3", "Listennummer 3")
    

    Dim oStyle As Style
    petrolLight = RGB(217, 233, 237)
    ockerLight = RGB(240, 232, 227)
    redLight = RGB(244, 226, 226)
    
    Call CreateOrEditTable("Scroll Tip", petrolLight)
    Call CreateOrEditTable("Scroll Info", ockerLight)
    Call CreateOrEditTable("Scroll Note", ockerLight)
    Call CreateOrEditTable("Scroll Warning", redLight)
    
End Sub

Sub CreateOrEditTable(styleName As String, color)
    Set oStyle = CreateOrEditTableStyle(styleName, "Normale Tabelle")
    oStyle.Table.Shading.BackgroundPatternColor = color
End Sub


Function CreateOrEditTableStyle(styleName As String, baseStyleName As String) As Style
    Dim doc As Document
    Set doc = Application.ActiveDocument
    
    Dim oStyle As Style
    If StyleExists(styleName) Then
        Debug.Print (doc.Styles(styleName))
        Set oStyle = doc.Styles(styleName)
    Else
        Set oStyle = doc.Styles.Add(name:=styleName, Type:=WdStyleType.wdStyleTypeTable)
    End If
    oStyle.baseStyle = baseStyleName
    Set CreateOrEditTableStyle = oStyle
End Function


Function CreateOrEditStyle(styleName As String, baseStyleName As String) As Style
    Dim doc As Document
    Set doc = Application.ActiveDocument
    
    Dim oStyle As Style
    If StyleExists(styleName) Then
        Set oStyle = doc.Styles(styleName)
    Else
        Set oStyle = doc.Styles.Add(styleName, WdStyleType.wdStyleTypeParagraphOnly)
    End If
    oStyle.baseStyle = baseStyleName
    Set CreateOrEditStyle = oStyle
End Function

Function StyleExists(styleName As String) As Boolean
    Dim oStyle As Style
    StyleExists = False
    For Each oStyle In ActiveDocument.Styles
        If oStyle.NameLocal = styleName Then
            StyleExists = True
            Exit Function
        End If
    Next oStyle
    Exit Function
End Function

