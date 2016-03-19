VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShieldingForm 
   Caption         =   "Shielding"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   -105
   ClientWidth     =   4920
   OleObjectBlob   =   "ShieldingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ShieldingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub selfButton_Click()
    Label.Caption = "thickness (cm)"
End Sub
Private Sub snowButton_Click()
    Label.Caption = "depth (cm) - dens (g/cm3) - ... - depth - dens"
End Sub
Private Sub topoButton_Click()
    Label.Caption = "strike - dip - azim - elev - ... - azim - elev"
End Sub
Private Sub userform_initialize()
    Me.topoButton.Value = True
    Label.Caption = "strike - dip - azim - elev - ... - azim - elev"
    RefEdit.Value = Selection.Address
End Sub
Private Sub OK_Click()
    On Error Resume Next
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.RefEdit.Value)
    If topoButton.Value Then
        Call topoShielding(theRange)
    ElseIf selfButton.Value Then
        Call selfShielding(theRange)
    ElseIf snowButton.Value Then
        Call snowShielding(theRange)
    End If
    Set theRange = Nothing
    Unload Me
End Sub
Private Sub snowShielding(theRange As MyRange)
If theRange.numcols >= 2 Then
    If Range(theRange.CellAddress(1, theRange.numcols)).row > 1 Then
       Range(theRange.CellAddress(1, theRange.numcols)).Offset(-1, 1).Value = "Shielding factor"
    End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        N = theRange.nonEmptyCols(rownum)
        num = N / 2
        S = 0
        For i = 0 To num - 1
            z = Range(theRange.CellAddress(rownum, 1 + i * 2)).Value
            rho = Range(theRange.CellAddress(rownum, 2 + i * 2)).Value
            S = S + (1 / num) * Exp(-z * rho / glob.L0)
        Next i
Label:
        Range(theRange.CellAddress(rownum, theRange.numcols)).Offset(0, 1) = S
        Range(theRange.CellAddress(rownum, theRange.numcols)).Offset(0, 1).NumberFormat = "0.00"
errorhandler:
        If Err.Number <> 0 Then
            Resume Label
        End If
    Next
Else
    MsgBox ("Please select at least two columns")
End If
End Sub
Private Sub selfShielding(theRange As MyRange)
    If Range(theRange.CellAddress(1, 1)).row > 1 Then
       Range(theRange.CellAddress(1, 1)).Offset(-1, 1).Value = "Shielding Factor"
    End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        z = theRange.CellValue(rownum, 1)
        If z = 0 Then
            Qs = 1
        Else
            Qs = (glob.L0 / (glob.rho * z)) * (1 - Exp(((-1 * glob.rho * z) / glob.L0)))
        End If
Label:
        Range(theRange.CellAddress(rownum, 1)).Offset(0, 1).Value = Qs
        Range(theRange.CellAddress(rownum, 1)).Offset(0, 1).NumberFormat = "0.00"
errorhandler:
        If Err.Number <> 0 Then
            Qs = 0
            Resume Label
        End If
    Next rownum
End Sub
Private Sub topoShielding(theRange As MyRange)
' this function was translated to VBA from Greg Balco's
' Matlab script skyline.m
Dim row As Range
Const B = (1 / 360) * (2 * pi)
Dim angles(360)
Dim a(360)
Dim horiz1(360)
Dim horiz2(360)
Dim horiz(360)

' initialize
For i = 0 To 360
    angles(i) = i * pi / 180
    a(i) = 0
    horiz1(i) = 0
    horiz2(i) = 0
    horiz(i) = 0
Next i
If theRange.numcols >= 2 Then
    If Range(theRange.CellAddress(1, theRange.numcols)).row > 1 Then
       Range(theRange.CellAddress(1, theRange.numcols)).Offset(-1, 1).Value = "Shielding factor"
    End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        ' read in the data
        strikeR = (theRange.CellValue(rownum, 1) / 360) * (2 * pi)
        dipR = (theRange.CellValue(rownum, 2) / 360) * (2 * pi)
        N = theRange.nonEmptyCols(rownum)
        numAz = (N - 2) / 2
        ReDim Az(numAz)
        ReDim AzR(numAz)
        ReDim AzR2(numAz + 2)
        ReDim El(numAz)
        ReDim ElR(numAz)
        ReDim ElR2(numAz + 2)
        strike = theRange.CellValue(rownum, 1)
        dip = theRange.CellValue(rownum, 2)
        For i = 0 To numAz - 1
            Az(i) = theRange.CellValue(rownum, 3 + 2 * i)
            El(i) = theRange.CellValue(rownum, 4 + 2 * i)
            AzR(i) = (Az(i) / 360) * 2 * pi
            ElR(i) = (El(i) / 360) * 2 * pi
            AzR2(i + 1) = AzR(i)
            ElR2(i + 1) = ElR(i)
        Next i
        S = 0
        If numAz > 0 Then ' if the user has entered azimuths and elevations
            AzR2(0) = AzR(numAz) - 2 * pi
            AzR2(numAz + 2) = AzR(0) + 2 * pi
            ElR2(0) = ElR(numAz)
            ElR2(numAz + 2) = ElR(0)
        End If
        For i = 0 To 360
            a(i) = angles(i) - (strikeR - (pi / 2))
            horiz1(i) = Atn(Tan(dipR) * Cos(a(i)))
            If horiz1(i) < 0 Then
                horiz1(i) = 0
            End If
            If numAz > 0 Then
                horiz2(i) = interp(angles(i), AzR2, ElR2)
                If horiz2(i) > horiz1(i) Then
                    horiz(i) = horiz2(i)
                Else
                    horiz(i) = horiz1(i)
                End If
            Else
                horiz(i) = horiz1(i)
            End If
            S = S + (B / (2 * pi)) * (Sin(horiz(i)) ^ (1 + glob.exponent))
        Next i
        topoFact = 1 - S
Label:
        Range(theRange.CellAddress(rownum, theRange.numcols)).Offset(0, 1) = topoFact
        Range(theRange.CellAddress(rownum, theRange.numcols)).Offset(0, 1).NumberFormat = "0.000"
errorhandler:
        If Err.Number <> 0 Then
            topoFact = 0
            Resume Label
        End If
    Next rownum
Else
    MsgBox ("Please select at least two columns")
End If
End Sub
Private Function interp(ByVal XI As Double, x As Variant, y As Variant) As Double
    Dim i As Integer
    For i = 1 To UBound(x) - 1
        If XI >= x(i) And XI <= x(i + 1) Then
            interp = y(i) + ((XI - x(i)) / (x(i + 1) - x(i))) * (y(i + 1) - y(i))
            Exit For
        End If
    Next i
End Function
Private Sub TopoHelpButton_Click()
    Dim Msg As String
    If topoButton.Value Then
        Msg = "" & vbCrLf
        Msg = Msg & "Topographic shielding: " & vbCrLf
        Msg = Msg & "-------------------------- " & vbCrLf
        Msg = Msg & "Select a range of cells, at least two columns wide. " & vbCrLf & vbCrLf
        Msg = Msg & "    s(1,1) d(1,1) a(1,1) e(1,1) a(1,2) e(1,2) ... a(1,N1) e(1,N1)    " & vbCrLf
        Msg = Msg & "    s(2,1) d(2,1) a(2,1) e(2,1) a(2,2) e(2,2) ... a(2,N2) e(2,N2)    " & vbCrLf
        Msg = Msg & "       :      :      :      :      :      :          :       :       " & vbCrLf
        Msg = Msg & "    s(n,1) d(n,1) a(n,1) e(n,1) a(n,2) e(n,2) ... a(n,Nn) e(n,Nn)    " & vbCrLf & vbCrLf
        Msg = Msg & "with s the strike and d the dip of the sample surface                " & vbCrLf
        Msg = Msg & "and  a and e a series of N azimuths and elevations                   " & vbCrLf
        Msg = Msg & "(it is not necessary for all rows to have the same number of columns)" & vbCrLf & vbCrLf
    ElseIf selfButton.Value Then
        Msg = Msg & ""
        Msg = Msg & "Self shielding (spallation only): " & vbCrLf
        Msg = Msg & "----------------------------------- " & vbCrLf
        Msg = Msg & "Select one column of cells." & vbCrLf & vbCrLf
        Msg = Msg & "    x(1)" & vbCrLf
        Msg = Msg & "     :  " & vbCrLf
        Msg = Msg & "    x(n)" & vbCrLf & vbCrLf
        Msg = Msg & "Where each x represents a sample thickness (in cm)." & vbCrLf
        Msg = Msg & "Sample density and spallogenic attenuation lengths can" & vbCrLf
        Msg = Msg & "be changed in the Settings menu ..." & vbCrLf & vbCrLf
    ElseIf snowButton.Value Then
        Msg = Msg & ""
        Msg = Msg & "Snow shielding: " & vbCrLf
        Msg = Msg & "----------------------- " & vbCrLf
        Msg = Msg & "Select a range of cells, at least two columns wide. " & vbCrLf & vbCrLf
        Msg = Msg & "    d(1,1) r(1,1) d(1,1) r(1,1) d(1,2) r(1,2) ... d(1,N1) r(1,N1)    " & vbCrLf
        Msg = Msg & "    d(2,1) r(2,1) d(2,1) r(2,1) d(2,2) r(2,2) ... d(2,N2) r(2,N2)    " & vbCrLf
        Msg = Msg & "       :      :      :      :      :      :          :       :       " & vbCrLf
        Msg = Msg & "    d(n,1) r(n,1) d(n,1) r(n,1) d(n,2) r(n,2) ... d(n,Nn) r(n,Nn)    " & vbCrLf & vbCrLf
        Msg = Msg & "with d the depth (in cm) and r the density (in g/cm3) of the snow    " & vbCrLf
        Msg = Msg & "(it is not necessary for all rows to have the same number of columns)" & vbCrLf & vbCrLf
    End If
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
