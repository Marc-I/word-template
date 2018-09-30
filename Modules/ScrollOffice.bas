Attribute VB_Name = "ScrollOffice"
    
Public Sub Replace(ByRef cc As contentControl, ByRef dic As Dictionary)
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
    
    Dim vT As Variant
    Set vT = dic.Item(t)
    Dim v As String
    v = CStr(vT)
 
    Dim r As Range
    Set r = cc.Range
        
    cc.Delete
    r.Delete
    r.InsertAfter (v)
    
End Sub
    
    
Sub ConvertToPageProperties()
    Dim doc As Document
    Set doc = Application.ActiveDocument
    
    Dim sr As Range
    For Each sr In doc.StoryRanges
        Dim cc As contentControl
        For Each cc In sr.ContentControls
            Call Replace(cc)
        Next
    Next
End Sub

Sub ConvertToPageProperties()
    

End Sub
