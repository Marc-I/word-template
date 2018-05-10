Attribute VB_Name = "XMLParts"
Public Function Load(ident As identification) As String
    Dim part As CustomXMLPart
    Dim node As CustomXMLNode

    With Application.ActiveDocument
        Set part = .CustomXMLParts(ident.namespace)
        Set node = part.SelectSingleNode("//*[name() = '" + ident.elementName + "']")
        Load = node.Text
    End With
End Function

Public Sub StoreDate(ident As identification, Value As Variant)
       Call StoreString(ident, Format(Value, "yyyy-mm-ddT00:00:00"))
End Sub

Public Sub StoreString(ident As identification, Value As String)
    Dim part As CustomXMLPart
    Dim node As CustomXMLNode

    With Application.ActiveDocument
        Set part = .CustomXMLParts(ident.namespace)
        Set node = part.SelectSingleNode("//*[name() = '" + ident.elementName + "']")
        node.Text = Value
    End With
End Sub

Public Sub DebugPrintNamespace(namespace As String)
    With Application.ActiveDocument
        Dim part As CustomXMLPart
        Set part = .CustomXMLParts(namespace)
        ' CostumXMLNodes: https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/customxmlnodes-object-office
        Dim nodes As CustomXMLNodes
        Set nodes = part.SelectNodes("//*")
        Dim index As Integer
        For index = 1 To nodes.Count
            Dim node As CustomXMLNode
            Debug.Print nodes.Item(index).BaseName
            Debug.Print nodes.Item(index).XML
        Next
    End With
End Sub
