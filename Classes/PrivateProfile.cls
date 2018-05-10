VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrivateProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private UserFile As String

Private Sub Class_Initialize()
    ' Grundeinstellungen - User
    UserPath = Options.DefaultFilePath(Path:=wdUserTemplatesPath)
    UserFile = UserPath & Application.PathSeparator & "htw_word.ini"
    
    ' Grundeinstellungen - Setup
    SetupPath = ActiveDocument.AttachedTemplate.Path & Application.PathSeparator
    SetupFile = SetupPath & "htwchur_setup.ini"
   
End Sub

Public Function Load(Key As String) As String
     Load = System.PrivateProfileString(UserFile, Key, Key)
End Function

Public Sub Store(Key As String, Value As String)
    System.PrivateProfileString(UserFile, Key, Key) = Value
End Sub


Public Function LoadDate(Key As String) As Date
On Error GoTo TypeError:
     Value = System.PrivateProfileString(UserFile, Key, Key)
     dValue = CDate(Value)
     LoadDate = dValue
     Exit Function

TypeError:
    LoadDate = Date
    
End Function

Public Sub StoreDate(Key As String, Value As Date)
    System.PrivateProfileString(UserFile, Key, Key) = Format(Value)
End Sub

' Format(Now(), "ddmmyy")
