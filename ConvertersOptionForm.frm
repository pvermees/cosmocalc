VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertersOptionForm 
   Caption         =   "Options"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   -540
   ClientWidth     =   3600
   OleObjectBlob   =   "ConvertersOptionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConvertersOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub userform_initialize()
    Me.replaceOptionButton.Value = glob.Replace
End Sub
Private Sub OKButton_Click()
    glob.Replace = Me.replaceOptionButton.Value
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
