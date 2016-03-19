VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BananaOptionForm 
   Caption         =   "Banana Options"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   -975
   ClientWidth     =   6600
   OleObjectBlob   =   "BananaOptionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BananaOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub userform_initialize()
    Me.xMinBox.Text = CStr(glob.xMin)
    Me.xMaxBox.Text = CStr(glob.xMax)
    Me.yMinBox.Text = CStr(glob.yMin)
    Me.yMaxBox.Text = CStr(glob.yMax)
    If glob.PlotEllipse Then
        Me.EllipseButton.Value = True
        Me.ErrorBarButton.Value = False
    Else
        Me.EllipseButton.Value = False
        Me.ErrorBarButton.Value = True
    End If
    If BananaForm.AlBeOptionButton.Value = True Then
        If glob.AlBeLogOrLin = "log" Then
            Me.logOption.Value = True
        ElseIf glob.AlBeLogOrLin = "lin" Then
            Me.linOption.Value = True
        End If
    ElseIf BananaForm.NeBeOptionButton.Value = True Then
        If glob.NeBeLogOrLin = "log" Then
            Me.logOption.Value = True
        ElseIf glob.NeBeLogOrLin = "lin" Then
            Me.linOption.Value = True
        End If
    End If
    Me.detailOption.Value = glob.detail
    Me.zeroErosionOption.Value = glob.zeroerosion
    Me.sigmaBox.Text = CStr(glob.sigma)
End Sub
Private Sub OKButton_Click()
    glob.xMin = CDbl(Me.xMinBox.Text)
    glob.xMax = CDbl(Me.xMaxBox.Text)
    glob.yMin = CDbl(Me.yMinBox.Text)
    glob.yMax = CDbl(Me.yMaxBox.Text)
    glob.PlotEllipse = Me.EllipseButton.Value
    If BananaForm.AlBeOptionButton.Value = True Then
        If Me.logOption.Value = True Then
            glob.AlBeLogOrLin = "log"
        ElseIf Me.linOption.Value = True Then
            glob.AlBeLogOrLin = "lin"
        End If
    ElseIf BananaForm.NeBeOptionButton.Value = True Then
        If Me.logOption.Value = True Then
            glob.NeBeLogOrLin = "log"
        ElseIf Me.linOption.Value = True Then
            glob.NeBeLogOrLin = "lin"
        End If
    End If
    glob.detail = Me.detailOption.Value
    glob.zeroerosion = Me.zeroErosionOption.Value
    glob.sigma = CDbl(Me.sigmaBox.Text)
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
