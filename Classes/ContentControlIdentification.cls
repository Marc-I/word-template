VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentControlIdentification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Connect data (identified by the id) with a content control (identified by the tag)
Private pId As String
Private pTag As String

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(id As String)
    pId = id
End Property

Public Property Get tag() As String
    tag = pTag
End Property

Public Property Let tag(tag As String)
    pTag = tag
End Property
    
Public Sub Initialize(id As String, tag As String)
    pId = id
    pTag = tag
End Sub


