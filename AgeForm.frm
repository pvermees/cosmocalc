VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AgeForm 
   Caption         =   "Age/Erosion rate calculator"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   -540
   ClientWidth     =   7440
   OleObjectBlob   =   "AgeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AgeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ageButton_Click()
    ComboBoxE.Visible = True
    LabelE1.Visible = True
    LabelE2.Visible = True
End Sub
Private Sub erosionButton_Click()
    ComboBoxE.Visible = False
    LabelE1.Visible = False
    LabelE2.Visible = False
End Sub

Private Sub userform_initialize()
    With Me.ComboBox
        .AddItem "26Al"
        .AddItem "10Be"
        .AddItem "21Ne"
        .AddItem "3He"
        .AddItem "36Cl"
        .AddItem "14C"
    End With
    With Me.ComboBox1
        .AddItem "26Al"
        .AddItem "10Be"
        .AddItem "21Ne"
        .AddItem "3He"
        .AddItem "36Cl"
        .AddItem "14C"
    End With
    With Me.ComboBox2
        .AddItem "26Al"
        .AddItem "10Be"
        .AddItem "21Ne"
        .AddItem "3He"
        .AddItem "36Cl"
        .AddItem "14C"
    End With
    With Me.ComboBox3
        .AddItem "26Al"
        .AddItem "10Be"
        .AddItem "21Ne"
        .AddItem "3He"
        .AddItem "36Cl"
        .AddItem "14C"
    End With
    With Me.ComboBox4
        .AddItem "Burial-Erosion"
        .AddItem "Burial-Exposure"
        .AddItem "Age-Erosion"
    End With
    singleNRefEdit.Value = Selection.Address
    twoNuclideRefEdit.Value = Selection.Address
    sBox.Text = CStr(1)
    tBox.Text = "inf"
    eBox.Text = CStr(0.1)
    tauBox.Text = CStr(1000)
End Sub
Private Sub forwardButton_Click()
    On Error Resume Next
    Dim nucl As MyNuclide
    Dim N As Double
    Dim E As Double
    Dim t As Variant
    Dim Tau As Double
    Set nucl = New MyNuclide
    Call nucl.SetProperties(ComboBox3.Value)
    E = CDbl(Me.eBox.Text)
    Tau = CDbl(Me.tauBox.Text)
    S = CDbl(Me.sBox.Text)
    Call nucl.SetScaling(S)
    If Me.tBox.Text = "inf" Then
        t = Me.tBox.Text
        N = getN(E / 1000, t, Tau * 1000, nucl)
    Else
        t = CDbl(Me.tBox.Text)
        N = getN(E / 1000, t * 1000, Tau * 1000, nucl)
    End If
    Me.NBox.Value = CLng(N)
    Set nucl = Nothing
End Sub
Private Sub oneNuclideButton_Click()
    On Error Resume Next
    Dim nucl As MyNuclide
    Set nucl = New MyNuclide
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.singleNRefEdit.Value)
    Call nucl.SetProperties(ComboBox.Value)
    If ageButton Then
        Call AgeCalc(nucl, theRange)
    ElseIf erosionButton Then
        Call ErosionCalc(nucl, theRange)
    End If
    Set nucl = Nothing
    Set theRange = Nothing
    Unload Me
End Sub
Private Sub twoNuclideButton_Click()
    If glob.NewtonOption Then
        Call twoNuclideCalc
    Else
        MetropolisForm.Show
    End If
    Unload Me
End Sub
Private Sub AgeCalc(nucl As MyNuclide, theRange As MyRange)
' single nuclide age calculation
        If theRange.numcols = 3 Then
            With Range(theRange.CellAddress(1, theRange.numcols))
            If .row > 1 Then
               .Offset(-1, 1).Value = "Age (ka)"
               .Offset(-1, 2).Value = "Err (1s)"
            End If
            End With
            For rownum = 1 To theRange.numRows
                On Error GoTo errorhandler
                Scaling = theRange.CellValue(rownum, 1)
                If Scaling <> "" And IsNumeric(Scaling) Then
                    Call nucl.SetScaling(Scaling)
                    N1 = theRange.CellValue(rownum, 2)
                    N1err = theRange.CellValue(rownum, 3)
                    E = CDbl(ComboBoxE.Text) / 1000
                    Age = getAge(N1, nucl, E) / 1000
                    AgeErr = getAgeErr(N1, N1err, nucl, Age, E) / 1000
Label:
                With Range(theRange.CellAddress(rownum, theRange.numcols))
                    .Offset(0, 1).Value = Age
                    .Offset(0, 2).Value = AgeErr
                    .Offset(0, 1).NumberFormat = "0.0"
                    .Offset(0, 2).NumberFormat = "0.0"
                End With
errorhandler:
                    If Err.Number <> 0 Then
                        Age = 0
                        AgeErr = 0
                        Resume Label
                    End If
                End If
            Next rownum
        Else
            MsgBox ("Please select three columns of data")
        End If
End Sub
Private Sub ErosionCalc(nucl As MyNuclide, theRange As MyRange)
        If theRange.numcols = 3 Then
            With Range(theRange.CellAddress(1, theRange.numcols))
            If .row > 1 Then
                .Offset(-1, 1).Value = "Erosion rate (cm/kyr)"
                .Offset(-1, 2).Value = "Err (1s)"
            End If
            End With
            For rownum = 1 To theRange.numRows
                On Error GoTo errorhandler
                Scaling = theRange.CellValue(rownum, 1)
                If Scaling <> "" And IsNumeric(Scaling) Then
                    Call nucl.SetScaling(Scaling)
                    N1 = theRange.CellValue(rownum, 2)
                    N1err = theRange.CellValue(rownum, 3)
                    erosion = 1000 * getErosion(N1, nucl)
                    ErosionErr = 1000 * getErosionErr(N1, N1err, nucl)
Label:
                    With Range(theRange.CellAddress(rownum, theRange.numcols))
                        .Offset(0, 1).Value = erosion
                        .Offset(0, 2).Value = ErosionErr
                        .Offset(0, 1).NumberFormat = "0.000"
                        .Offset(0, 2).NumberFormat = "0.000"
                    End With
errorhandler:
                    If Err.Number <> 0 Then
                        erosion = 0
                        ErosionErr = 0
                        Resume Label
                    End If
                End If
            Next rownum
        Else
            MsgBox ("Please select four columns of data")
        End If
End Sub
Private Sub twoNuclideOptionButton_Click()
    Call twoNuclideOptionForm.Show
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub CommandButton2_Click()
    Unload Me
End Sub
Private Sub CommandButton3_Click()
    Unload Me
End Sub
Private Sub twoNuclideCalc()
    Dim opt As String
    Dim S1 As Variant
    Dim nucl1 As MyNuclide
    Dim nucl2 As MyNuclide
    Dim x As Double, y As Double, Xerr As Double, Yerr As Double
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.twoNuclideRefEdit.Value)
    Set nucl1 = New MyNuclide
    Set nucl2 = New MyNuclide
    Call nucl1.SetProperties(AgeForm.ComboBox1.Value)
    Call nucl2.SetProperties(AgeForm.ComboBox2.Value)
    
    If theRange.numcols <> 6 Then
        MsgBox ("Please select six columns of data")
    Else
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler:
        S1 = theRange.CellValue(rownum, 1)
        If S1 <> "" And IsNumeric(S1) Then
            Call nucl1.SetScaling(S1)
            N1 = theRange.CellValue(rownum, 2)
            N1err = theRange.CellValue(rownum, 3)
            S2 = theRange.CellValue(rownum, 4)
            Call nucl2.SetScaling(S2)
            N2 = theRange.CellValue(rownum, 5)
            N2err = theRange.CellValue(rownum, 6)
            If Me.ComboBox4.Value = "Age-Erosion" Then
                ' two nuclide exposure age (Y) - erosion rate X) calculation
                Call getAgeErosion(N1, N1err, N2, N2err, nucl1, nucl2, x, Xerr, y, Yerr)
            ElseIf Me.ComboBox4.Value = "Burial-Erosion" Then
                ' two nuclide burial age (Y) - erosion rate (X) calculation
                Call getBurialErosion(N1, N1err, N2, N2err, nucl1, nucl2, x, Xerr, y, Yerr)
            ElseIf Me.ComboBox4.Value = "Burial-Exposure" Then
                ' two nuclide burial age (Y) - exposure age (X) calculation
                Call getBurialExposure(N1, N1err, N2, N2err, nucl1, nucl2, x, Xerr, y, Yerr)
            End If
Label:
            With Range(theRange.CellAddress(rownum, theRange.numcols))
                .Offset(0, 1).Value = y / 1000
                .Offset(0, 1).NumberFormat = "0.0"
                .Offset(0, 2).Value = Yerr / 1000
                .Offset(0, 2).NumberFormat = "0.0"
                If Me.ComboBox4.Value = "Burial-Exposure" Then
                    .Offset(0, 3).Value = x / 1000
                    .Offset(0, 3).NumberFormat = "0.0"
                    .Offset(0, 4).Value = Xerr / 1000
                    .Offset(0, 4).NumberFormat = "0.0"
                Else
                    .Offset(0, 3).Value = x * 1000
                    .Offset(0, 3).NumberFormat = "0.000"
                    .Offset(0, 4).Value = Xerr * 1000
                    .Offset(0, 4).NumberFormat = "0.000"
                End If
            End With
errorhandler:
            If (Err.Number <> 0) Then
                'MsgBox ("Newton method did not converge. Try again or use Metropolis option.")
                x = 0
                Xerr = 0
                y = 0
                Yerr = 0
                Resume Label
            End If
        End If
    Next rownum
    'print header:
    If Range(theRange.CellAddress(1, theRange.numcols)).row > 1 Then
        With Range(theRange.CellAddress(1, theRange.numcols))
            If AgeForm.ComboBox4.Value = "Age-Erosion" Then
                .Offset(-1, 1).Value = "Exposure Age (ka)"
                .Offset(-1, 3).Value = "Erosion Rate (cm/ka)"
            ElseIf AgeForm.ComboBox4.Value = "Burial-Erosion" Then
                .Offset(-1, 1).Value = "Burial Age (ka)"
                .Offset(-1, 3).Value = "Erosion Rate (cm/ka)"
            ElseIf AgeForm.ComboBox4.Value = "Burial-Exposure" Then
                .Offset(-1, 1).Value = "Burial Age (ka)"
                .Offset(-1, 3).Value = "Exposure Age (ka)"
            End If
            .Offset(-1, 2).Value = "1 sigma"
            .Offset(-1, 4).Value = "1 sigma"
        End With
    End If
    End If
    Set theRange = Nothing
    Set nucl1 = Nothing
    Set nucl2 = Nothing
    Unload Me
End Sub
Private Sub twoNuclideHelpButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Select a range of cells, six columns wide " & vbCrLf & vbCrLf
    Msg = Msg & "        S1(1) N1(1) N1err(1) S2(1) N2(1) N2err(1)" & vbCrLf
    Msg = Msg & "             :      :            :        :         :            :     " & vbCrLf
    Msg = Msg & "        S1(n) N1(n) N1err(n) S2(n) N2(n) N2err(n)" & vbCrLf & vbCrLf
    Msg = Msg & "with:         " & vbCrLf
    Msg = Msg & "   - S1 = Composite scaling factor for Nuclide 1, i.e." & vbCrLf
    Msg = Msg & "          the combined effects of latitude, elevation," & vbCrLf
    Msg = Msg & "          snow- and self-shielding" & vbCrLf
    Msg = Msg & "   - N1 = Concentration of Nuclide 1 (atoms/g)" & vbCrLf
    Msg = Msg & "   - N1err = 1-sigma measurement uncertainty of N1" & vbCrLf
    Msg = Msg & "   - S2 = Scaling factor of Nuclide 2" & vbCrLf
    Msg = Msg & "   - N2 = Concentration of Nuclide 2 (atoms/g)" & vbCrLf
    Msg = Msg & "   - N2err = 1-sigma measurement uncertainty of N2" & vbCrLf & vbCrLf
    Msg = Msg & "If the output consists of only zeros, try using" & vbCrLf
    Msg = Msg & "the Metropolis algorithm instead of Newton's method" & vbCrLf
    Msg = Msg & "(see 'Options' menu)" & vbCrLf & vbCrLf
    Msg = Msg & "N.B: N1 and N2 must be topography-corrected, unless" & vbCrLf
    Msg = Msg & "     the topographic shielding correction is small" & vbCrLf
    Msg = Msg & "     (S_t > 0.95)" & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub ForwardHelpButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "The forward modeling tool calculates the TCN" & vbCrLf
    Msg = Msg & "concentration expected for a given nuclide, " & vbCrLf
    Msg = Msg & "scaling + shielding factor, exposure age, " & vbCrLf
    Msg = Msg & "erosion rate and/or burial age. To model " & vbCrLf
    Msg = Msg & "steady state conditions, enter 'inf'" & vbCrLf
    Msg = Msg & "(without apostrophes) as the exposure age" & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub ErosionHelpButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Select a range of cells, three columns wide " & vbCrLf & vbCrLf
    Msg = Msg & "          S(1) N(1) Nerr(1)          " & vbCrLf
    Msg = Msg & "               :     :      :              " & vbCrLf
    Msg = Msg & "          S(n) N(n) Nerr(n)          " & vbCrLf & vbCrLf
    Msg = Msg & "with:         " & vbCrLf
    Msg = Msg & "   - S = Composite scaling factor; the product of latitude," & vbCrLf
    Msg = Msg & "         elevation, snow- and self-shielding corrections." & vbCrLf
    Msg = Msg & "   - N = Nuclide concentration (atoms/g)" & vbCrLf
    Msg = Msg & "   - Nerr = 1-sigma measurement uncertainty of N" & vbCrLf & vbCrLf
    Msg = Msg & "N.B: N must be topography-corrected, unless the" & vbCrLf
    Msg = Msg & "     topographic shielding correction is small (S_t > 0.95)." & vbCrLf
    Msg = Msg & "     In that case, the topographic shielding correction" & vbCrLf
    Msg = Msg & "     can be lumped together with S." & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
