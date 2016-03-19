VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MetropolisForm 
   Caption         =   "Progress"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   -540
   ClientWidth     =   4320
   OleObjectBlob   =   "MetropolisForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MetropolisForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub userform_activate()
    On Error Resume Next
    Dim nucl1 As MyNuclide
    Dim nucl2 As MyNuclide
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(AgeForm.twoNuclideRefEdit.Value)
    Set nucl1 = New MyNuclide
    Set nucl2 = New MyNuclide
    Call nucl1.SetProperties(AgeForm.ComboBox1.Value)
    Call nucl2.SetProperties(AgeForm.ComboBox2.Value)
    If theRange.numcols = 6 Then
    With Range(theRange.CellAddress(1, theRange.numcols))
        If AgeForm.ComboBox4.Value = "Age-Erosion" Then
        ' two nuclide age - erosion calculation
            Call MetroCalc(theRange, "age_erosion", nucl1, nucl2)
            .Offset(-1, 1).Value = "Exposure Age (ka)"
            .Offset(-1, 4).Value = "Erosion (cm/ka)"
        ElseIf AgeForm.ComboBox4.Value = "Burial-Erosion" Then
        ' two nuclide burial age - erosion calculation
            Call MetroCalc(theRange, "burial_erosion", nucl1, nucl2)
            .Offset(-1, 1).Value = "Burial age (ka)"
            .Offset(-1, 4).Value = "Erosion (cm/ka)"
        ElseIf AgeForm.ComboBox4.Value = "Burial-Exposure" Then
        ' two nuclide burial age - exposure calculation
            Call MetroCalc(theRange, "burial_exposure", nucl1, nucl2)
            .Offset(-1, 1).Value = "Burial age (ka)"
            .Offset(-1, 4).Value = "Exposure age (ka)"
        End If
        .Offset(-1, 2).Value = glob.ConfiLevel / 2 & " pctile"
        .Offset(-1, 3).Value = 100 - glob.ConfiLevel / 2 & " pctile"
        .Offset(-1, 5).Value = glob.ConfiLevel / 2 & " pctile"
        .Offset(-1, 6).Value = 100 - glob.ConfiLevel / 2 & " pctile"
    End With
    Else
        MsgBox ("Please select six columns")
    End If
    Set theRange = Nothing
    Set nucl1 = Nothing
    Set nucl2 = Nothing
    Unload Me
End Sub
Private Sub MetroCalc(theRange As MyRange, ByVal opt As String, nuclide1 As MyNuclide, nuclide2 As MyNuclide)
' opt = "age_erosion", "burial_erosion" or "burial_exposure"
' X,Y = t,E; B,E or B,t
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        If opt = "age_erosion" Then
            sht = twoNuclideCalc(theRange, rownum, "age_erosion", nuclide1, nuclide2)
        ElseIf opt = "burial_erosion" Then
            sht = twoNuclideCalc(theRange, rownum, "burial_erosion", nuclide1, nuclide2)
        ElseIf opt = "burial_exposure" Then
            sht = twoNuclideCalc(theRange, rownum, "burial_exposure", nuclide1, nuclide2)
        End If
        If sht <> "" Then
            On Error GoTo errorhandler:
            ' if you want to plot the posterior distribution, you should uncomment the following line,
            ' the nine lines in TwoNuclideCalc and place a breakpoint here:
            ' Worksheets(sht).Visible = True
            Cell1 = "A" & Int(0.1 * glob.MetropIter)
            Cell2 = "C" & glob.MetropIter
            Worksheets(sht).Range(Cell1, Cell2).name = "wholeRange"
            Worksheets(sht).Range("wholeRange").Sort Key1:=Worksheets(sht).Range("A1"), Order1:=xlAscending, Header:=xlGuess, _
                OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
            Cell1 = "B" & Int((0.1 + glob.ConfiLevel / 100) * glob.MetropIter)
            Cell2 = "B" & glob.MetropIter
            Worksheets(sht).Range(Cell1, Cell2).name = "bestXrange"
            X05 = Application.min(Worksheets(sht).Range("bestXrange"))
            X95 = Application.Max(Worksheets(sht).Range("bestXrange"))
            Xbest = Worksheets(sht).Cells(glob.MetropIter, 2).Value
            Cell1 = "C" & Int((0.1 + glob.ConfiLevel / 100) * glob.MetropIter)
            Cell2 = "C" & glob.MetropIter
            Worksheets(sht).Range(Cell1, Cell2).name = "bestYrange"
            Y05 = Application.min(Worksheets(sht).Range("bestYrange"))
            Y95 = Application.Max(Worksheets(sht).Range("bestYrange"))
            Ybest = Worksheets(sht).Cells(glob.MetropIter, 3).Value
Label:
            With Range(theRange.CellAddress(rownum, theRange.numcols))
                .Offset(0, 1).Value = Xbest / 1000
                .Offset(0, 2).Value = X05 / 1000
                .Offset(0, 3).Value = X95 / 1000
                .Offset(0, 1).NumberFormat = "0.0"
                .Offset(0, 2).NumberFormat = "0.0"
                .Offset(0, 3).NumberFormat = "0.0"
                If opt = "burial_exposure" Then
                    .Offset(0, 4).Value = Ybest / 1000
                    .Offset(0, 5).Value = Y05 / 1000
                    .Offset(0, 6).Value = Y95 / 1000
                    .Offset(0, 4).NumberFormat = "0.0"
                    .Offset(0, 5).NumberFormat = "0.0"
                    .Offset(0, 6).NumberFormat = "0.0"
                Else
                    .Offset(0, 4).Value = Ybest * 1000
                    .Offset(0, 5).Value = Y05 * 1000
                    .Offset(0, 6).Value = Y95 * 1000
                    .Offset(0, 4).NumberFormat = "0.000"
                    .Offset(0, 5).NumberFormat = "0.000"
                    .Offset(0, 6).NumberFormat = "0.000"
                End If
            End With
            Call DeleteRangeName("wholeRange")
            Call DeleteRangeName("bestXrange")
            Call DeleteRangeName("bestYrange")
            Call removeSheet(sht)
errorhandler:
            If Err.Number <> 0 Then
                'MsgBox ("Error -- no solution found.")
                Xbest = 0
                Ybest = 0
                X05 = 0
                X95 = 0
                Y05 = 0
                Y95 = 0
                Resume Label
            End If
        End If
    Next rownum
End Sub
Private Function twoNuclideCalc(theRange As MyRange, ByVal rownum As Integer, ByVal opt As String, nuclide1 As MyNuclide, nuclide2 As MyNuclide) As String
' two nuclide age-erosion calculation using the Metropolis algorithm
    Dim Rnge As Range, sht As Worksheet, Yold As Double, Xold As Double, Lold As Double, S1 As Variant


    minX = 100 'yr
    maxX = 20000000 'yr
    If opt = "burial_exposure" Then
        minY = 100 'yr
        maxY = 20000000 'yr
    Else
        minY = glob.Zero 'cm/yr
        maxY = 0.1 'cm/yr
    End If
    mu = glob.rho / glob.L0
    '
    S1 = theRange.CellValue(rownum, 1)
    If S1 <> "" And IsNumeric(S1) Then
        Call nuclide1.SetScaling(S1)
        N1 = theRange.CellValue(rownum, 2)
        N1err = theRange.CellValue(rownum, 3)
        S2 = theRange.CellValue(rownum, 4)
        Call nuclide2.SetScaling(S2)
        N2 = theRange.CellValue(rownum, 5)
        N2err = theRange.CellValue(rownum, 6)
    
        ' initial guesses ...
        Call initialize(opt, nuclide1, nuclide2, N1, N2, N1err, N2err, Yold, Xold, Lold)
        Ynew = Yold
        Xnew = Xold
        Lratio = 0
        randNum = 0
        numsol = glob.MetropIter
        
        Call removeSheet("TempSheet")
        Set sht = Worksheets.Add
        sht.name = "TempSheet"
        sht.Visible = False
        i = 1 ' number of accepted solutions
        j = 1
        If glob.PC Then
            stepsize = 1 ' update the progress bar only every step%
        Else
            stepsize = 5 ' Excel is slower on the Mac
        End If
        numsteps = 1
        Do ' loop through the Metropolis iterations - hope nothing goes wrong since there's no error capturing for speed
            
            logYnew = randChange(0.95 * Log(Yold), 1.05 * Log(Yold), Log(minY), Log(maxY))
            Ynew = Exp(logYnew)
            logXnew = randChange(0.95 * Log(Xold), 1.05 * Log(Xold), Log(minX), Log(maxX))
            Xnew = Exp(logXnew)
            If opt = "age_erosion" Then
                Lnew = getAgeErosionL(Ynew, Xnew, nuclide1, N1, N1err, nuclide2, N2, N2err)
            ElseIf opt = "burial_erosion" Then
                Lnew = getBurialAgeErosionL(Ynew, Xnew, nuclide1, N1, N1err, nuclide2, N2, N2err)
            ElseIf opt = "burial_exposure" Then
                Lnew = getBurialExposureL(Ynew, Xnew, nuclide1, N1, N1err, nuclide2, N2, N2err)
            End If
            If Lnew >= Lold Then
                sht.Cells(i, 1) = Lnew
                sht.Cells(i, 2) = Xnew
                sht.Cells(i, 3) = Ynew
                ' the following 9 lines are for testing only and can be commented out
                'If opt = "age_erosion" Then
                '    N1Est = getN(Ynew, Tnew, 0, nuclide1)
                '    N2Est = getN(Ynew, Tnew, 0, nuclide2)
                'ElseIf opt = "burial_erosion" Then
                '    N1Est = getN(Ynew, "inf", Tnew, nuclide1)
                '    N2Est = getN(Ynew, "inf", Tnew, nuclide2)
                'End If
                'sht.Cells(i, 4) = N2Est
                'sht.Cells(i, 5) = N1Est / N2Est

                Lold = Lnew
                Xold = Xnew
                Yold = Ynew
                i = i + 1
                numsteps = updateBar(i / numsol, numsteps, stepsize)
            Else
                P = Exp(Lnew - Lold)
                randNum = Rnd
                If randNum <= P Then
                    sht.Cells(i, 1) = Lnew
                    sht.Cells(i, 2) = Xnew
                    sht.Cells(i, 3) = Ynew
                    ' the following 9 lines are for testing only and can be commented out
                    'If opt = "age_erosion" Then
                    '    N1Est = getN(Ynew, Xnew, 0, nuclide1)
                    '    N2Est = getN(Ynew, Xnew, 0, nuclide2)
                    'ElseIf opt = "burial_erosion" Then
                    '    N1Est = getN(Ynew, "inf", Xnew, nuclide1)
                    '    N2Est = getN(Ynew, "inf", Xnew, nuclide2)
                    'End If
                    'sht.Cells(i, 4) = N2Est
                    'sht.Cells(i, 5) = N1Est / N2Est
                    
                    Lold = Lnew
                    Xold = Xnew
                    Yold = Ynew
                    i = i + 1
                    numsteps = updateBar(i / numsol, numsteps, stepsize)
                End If
            End If
            j = j + 1
            If j > numsol * 200 Then
                MsgBox ("Iteration did not converge.")
                Exit Do
            End If
        Loop While i <= numsol
        twoNuclideCalc = sht.name
    Else
        twoNuclideCalc = ""
    End If
    Set sht = Nothing
End Function
Private Sub initialize(ByVal opt As String, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByVal N1 As Double, ByVal N2 As Double, ByVal N1err As Double, ByVal N2err As Double, ByRef y As Double, ByRef x As Double, ByRef L As Double)
    Dim Yerr As Double, Xerr As Double
    If opt = "age_erosion" Then
    '    Call getAgeErosion(N1, N1err, N2, N2err, nuclide1, nuclide2, Y, Yerr, X, Xerr)
        If nuclide1.L < nuclide2.L Then
            y = getErosion(N1, nuclide1)
            x = getAge(N1, nuclide1, y)
        Else
            y = getErosion(N2, nuclide2)
            x = getAge(N2, nuclide2, y)
        End If
        L = getAgeErosionL(y, x, nuclide1, N1, N1err, nuclide2, N2, N2err)
    ElseIf opt = "burial_erosion" Then
    '    Call getBurialErosion(N1, N1err, N2, N2err, nuclide1, nuclide2, Y, Yerr, X, Xerr)
        If nuclide1.L > nuclide2.L Then
            y = getErosion(N1, nuclide1)
            x = getBurial(N1, nuclide1, y / 2, "inf")
        Else
            y = getErosion(N2, nuclide2)
            x = getBurial(N2, nuclide2, y / 2, "inf")
        End If
        L = getBurialAgeErosionL(y, x, nuclide1, N1, N1err, nuclide2, N2, N2err)
    ElseIf opt = "burial_exposure" Then
    '    Call getBurialErosion(N1, N1err, N2, N2err, nuclide1, nuclide2, Y, Yerr, X, Xerr)
        If nuclide1.L > nuclide2.L Then
            y = getAge(N1, nuclide1, 0)
            x = getBurial(N1, nuclide1, 0, 2 * y)
        Else
            y = getAge(N2, nuclide2, 0)
            x = getBurial(N2, nuclide2, 0, 2 * y)
        End If
        L = getBurialAgeErosionL(y, x, nuclide1, N1, N1err, nuclide2, N2, N2err)
    End If
End Sub
Private Function updateBar(ByVal fractDone As Double, ByVal numsteps As Integer, ByVal stepsize As Double) As Integer
    If (100 * fractDone > numsteps * stepsize) Then
        With Me
            .FrameProgress.Caption = Format(fractDone, "0%")
            .LabelProgress.width = fractDone * (.FrameProgress.width - 10)
        End With
        DoEvents
        updateBar = numsteps + 1
    Else
        updateBar = numsteps
    End If
End Function
Private Function getAgeErosionL(ByVal E As Double, ByVal t As Double, _
                      nuclide1 As MyNuclide, ByVal N1 As Double, ByVal N1err As Double, _
                      nuclide2 As MyNuclide, ByVal N2 As Double, ByVal N2err As Double)
    N1est = getN(E, t, 0, nuclide1)
    N2est = getN(E, t, 0, nuclide2)
    getAgeErosionL = -Log(2 * pi * N1err * N2err) - 0.5 * _
                (((N1est - N1) ^ 2) / (N1err ^ 2) + ((N2est - N2) ^ 2) / (N2err ^ 2))
End Function
Private Function getBurialAgeErosionL(ByVal E As Double, ByVal t As Double, _
                      nuclide1 As MyNuclide, ByVal N1 As Double, ByVal N1err As Double, _
                      nuclide2 As MyNuclide, ByVal N2 As Double, ByVal N2err As Double)
    N1est = getN(E, "inf", t, nuclide1)
    N2est = getN(E, "inf", t, nuclide2)
    getBurialAgeErosionL = -Log(2 * pi * N1err * N2err) - 0.5 * _
                (((N1est - N1) ^ 2) / (N1err ^ 2) + ((N2est - N2) ^ 2) / (N2err ^ 2))
End Function
Private Function getBurialExposureL(ByVal t As Double, ByVal B As Double, _
                      nuclide1 As MyNuclide, ByVal N1 As Double, ByVal N1err As Double, _
                      nuclide2 As MyNuclide, ByVal N2 As Double, ByVal N2err As Double)
    N1est = getN(0, t, B, nuclide1)
    N2est = getN(0, t, B, nuclide2)
    getBurialExposureL = -Log(2 * pi * N1err * N2err) - 0.5 * _
                (((N1est - N1) ^ 2) / (N1err ^ 2) + ((N2est - N2) ^ 2) / (N2err ^ 2))
End Function
Private Function randChange(ByVal Low As Double, ByVal High As Double, ByVal absMin As Double, ByVal absMax As Double)
    If Low < absMin Then
        Low = absMin
    End If
    If High > absMax Then
        High = absMax
    End If
    randChange = Rnd * (High - Low) + Low
End Function
