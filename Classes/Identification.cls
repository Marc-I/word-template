VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Identification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Connect data (identified by the id) with a an XML component (identified by the element name and namespace)

Private pId As String
Private pElementName As String
Private pNamespace As String

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(id As String)
    pId = id
End Property

Public Property Get elementName() As String
    elementName = pElementName
End Property

Public Property Let elementName(elementName As String)
    pElementName = elementName
End Property
    
Public Property Get namespace() As String
    namespace = pNamespace
End Property

Public Property Let namespace(namespace As String)
    pNamespace = namespace
End Property

Public Sub Initialize(id As String, elementName As String, namespace As String)
    pId = id
    pElementName = elementName
    pNamespace = namespace
End Sub

