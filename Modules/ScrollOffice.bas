Attribute VB_Name = "ScrollOffice"
    
Public Sub Replace(ByRef cc As contentControl)
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
  
    Dim tV As Variant
    tV = cc.tag
    
    Dim t As String
    t = CStr(tV)
    Debug.Print (t)
    
    Dim v As String
    v = dic.Item(t)
 
    Dim r As Range
    Set r = cc.Range
        
    cc.Delete
    r.Delete
    r.InsertAfter (v)
    
End Sub
    
    
Sub ReplaceContentControls()
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
                Call Replace(cc)
            Next
            'Get next linked story (if any)
            Set rngStory = rngStory.NextStoryRange
        Loop Until rngStory Is Nothing
    Next
End Sub
    
Sub ConvertToPageProperties()
    Call ReplaceContentControls
    ' Deleting while Iterating seems a problem. Some content controls stay. So just do it several times
    Call ReplaceContentControls
    Call ReplaceContentControls
    
    '
    Call Selection.GoTo(wdGoToPage, wdGoToAbsolute, 3)
    Call Selection.EndKey(wdStory, wdExtend)
    Dim r As Range
    Set r = Selection.Range
    r.Delete
    r.InsertAfter ("$scroll.content")
  
End Sub

Sub ConvertToConfluence()
    

End Sub
