VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValuePair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copied from https://github.com/zrisher/vba_collection
'
' The VBA class Collection (https://msdn.microsoft.com/en-us/library/aa231021(v=vs.60).aspx) supports keys.
' But it has some disadvantages https://stackoverflow.com/questions/5702362/vba-collection-list-of-keys).

'Unrestricted class just to hold pairs of values together and permit Dictionary object updating
Public Key As String
Public Value As Variant

