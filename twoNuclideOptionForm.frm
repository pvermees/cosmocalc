VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} twoNuclideOptionForm 
   Caption         =   "Options"
   ClientHeight    =   2560
   ClientLeft      =   45
   ClientTop       =   -105
   ClientWidth     =   4440
   OleObjectBlob   =   "twoNuclideOptionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "twoNuclideOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub userform_initialize()
    Me.confiLevelBox.Value = 100 - glob.ConfiLevel
    If glob.NewtonOption Then
        Me.NewtonButton.Value = True
        Me.MetropolisButton.Value = False
    Else
        Me.NewtonButton.Value = False
        Me.MetropolisButton.Value = True
    End If
End Sub
Private Sub OKButton_Click()
    glob.ConfiLevel = 100 - CDbl(Me.confiLevelBox.Text)
    glob.NewtonOption = Me.NewtonButton.Value
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub HelpButton_Click()

End Sub
