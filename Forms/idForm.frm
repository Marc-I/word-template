VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} idForm 
   Caption         =   "Identifikation"
   ClientHeight    =   4920
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5028
   OleObjectBlob   =   "idForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "idForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public cancelled As Boolean

Private Sub cancelButton_Click()
    cancelled = True
    idForm.Hide
End Sub


Private Sub okButton_Click()
    cancelled = False
    idForm.Hide
End Sub
