VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BananaForm 
   Caption         =   "Banana Plots"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   -2745
   ClientWidth     =   5895
   OleObjectBlob   =   "BananaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BananaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub userform_initialize()
    Me.AlBeOptionButton.Value = True
    AlBeNeRefEdit.Value = Selection.Address
End Sub
Private Sub AlBeOptionButton_Change()
    glob.xMin = 10000
    glob.xMax = 100000000
    glob.yMin = 1
    glob.yMax = 9
End Sub
Private Sub NeBeOptionButton_Change()
    glob.xMin = 10000
    glob.xMax = 14000000
    glob.yMin = 3
    glob.yMax = 25
End Sub
Private Sub OptionButton_Click()
    Call BananaOptionForm.Show
End Sub
Private Sub PlotAlNeBeBanana_Click()
    On Error Resume Next
    Dim theRange As MyRange
    Dim msgtext As String
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.AlBeNeRefEdit.Value)
    If (theRange.numcols = 6) Or (theRange.numcols = 7) Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        If AlBeOptionButton.Value = True Then
            logorlin = glob.AlBeLogOrLin
            Call PlotAlBeBanana(logorlin)
        ElseIf NeBeOptionButton.Value = True Then
            logorlin = glob.NeBeLogOrLin
            Call PlotNeBeBanana(logorlin)
        End If
        Call DataPlot(theSheet, theRange, logorlin)
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    Else
        MsgBox "Please select six or seven columns"
    End If
    ActiveChart.Deselect
    Set theRange = Nothing
    Unload Me
End Sub
' plot the background-lines
Private Sub PlotAlBeBanana(ByVal logorlin As String)
    Dim nucl1 As MyNuclide
    Set nucl1 = New MyNuclide
    Dim nucl2 As MyNuclide
    Set nucl2 = New MyNuclide
    Call nucl1.SetProperties("26Al")
    Call nucl2.SetProperties("10Be")
    N = 12
    Dim tickArr(4)
    tickArr(1) = 2
    tickArr(2) = 4
    tickArr(3) = 6
    tickArr(4) = 7
    '
    If logorlin = "lin" Then
        Call XYPlot(nucl1.name, nucl2.name, logorlin)
    Else
        tickFormula = getTickFormula(tickArr)
        Call XYPlot(nucl1.name, nucl2.name, tickFormula)
        Call addTickLabels(glob.xMin, tickArr)
    End If
    E1 = glob.Zero
    E2 = 0.0001
    E3 = 10
    t1 = 1
    t2 = 10000
    t3 = 10000000
    ' Plot saturation line
    Call PlotVariableErosion(N, "inf", 0, E1, E2, nucl1, nucl2, 1, 2, logorlin)
    Call PlotVariableErosion(N, "inf", 0, E2, E3, nucl1, nucl2, 1, 2, logorlin)
    '
    ' Plot zero erosion line
    Call PlotVariableTime(N, 0, 0, t1, t2, nucl1, nucl2, 1, 2, logorlin)
    Call PlotVariableTime(N, 0, 0, t2, t3, nucl1, nucl2, 1, 2, logorlin)
    '
    ' Plot burial decay line after saturation
    Call PlotVariableBurial(N, "inf", 0, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
    '
    If glob.detail Then
        '
        If glob.zeroerosion Then
            '
            ' Plot 2Ma burial line
            Call PlotVariableTime(N, 0, 2000000, t1, t2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0, 2000000, t2, t3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 1Ma burial line
            Call PlotVariableTime(N, 0, 1000000, t1, t2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0, 1000000, t2, t3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 20ka exposure
            Call PlotVariableBurial(N, 20000, 0, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 100ka exposure
            Call PlotVariableBurial(N, 100000, 0, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 500ka exposure
            Call PlotVariableBurial(N, 500000, 0, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 2Ma exposure
            Call PlotVariableBurial(N, 2000000, 0, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            Call getNaddLabel(20000, 0, 100000, "20ka", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(100000, 0, 100000, "100ka", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(500000, 0, 100000, "500ka", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(2000000, 0, 100000, "2Ma", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(5000000, 0, 1000000, "1Ma burial", nucl1, nucl2, 10, logorlin)
            Call getNaddLabel(5000000, 0, 2000000, "2Ma burial", nucl1, nucl2, 10, logorlin)
            '
        Else
            '
            ' Plot 1m/Ma erosion line
            Call PlotVariableTime(N, 0.0001, 0, t1, t2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.0001, 0, t2, t3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 10m/Ma erosion line
            Call PlotVariableTime(N, 0.001, 0, t1, t2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.001, 0, t2, t3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 100m/Ma erosion line
            Call PlotVariableTime(N, 0.01, 0, t1, t2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.01, 0, t2, t3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 1Ma burial line
            Call PlotVariableErosion(N, "inf", 1000000, E1, E2, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 1000000, E2, E3, nucl1, nucl2, 10, 1, logorlin)
            '
            ' Plot 2Ma burial line
            Call PlotVariableErosion(N, "inf", 2000000, E1, E2, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 2000000, E2, E3, nucl1, nucl2, 10, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 1m/Ma
            Call PlotVariableBurial(N, "inf", 0.0001, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 10m/Ma
            Call PlotVariableBurial(N, "inf", 0.001, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 100m/Ma
            Call PlotVariableBurial(N, "inf", 0.01, 1, 10000000, nucl1, nucl2, 3, 1, logorlin)
            '
            Call getNaddLabel("inf", 0, 500000, "0cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel("inf", 0.0001, 500000, "0.1cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel("inf", 0.001, 500000, "1cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel("inf", 0.01, 500000, "10cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel("inf", 0.00001, 1000000, "1Ma burial", nucl1, nucl2, 10, logorlin)
            Call getNaddLabel("inf", 0.00001, 2000000, "2Ma burial", nucl1, nucl2, 10, logorlin)
            Call getNaddLabel(1000000, 0, 0, "1Ma", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(2000000, 0, 0, "2Ma", nucl1, nucl2, 7, logorlin)
            '
        End If
        
    End If
    Set nucl1 = Nothing
    Set nucl2 = Nothing
End Sub
Private Sub PlotNeBeBanana(ByVal logorlin As String)
    Dim nucl1 As MyNuclide
    Set nucl1 = New MyNuclide
    Dim nucl2 As MyNuclide
    Set nucl2 = New MyNuclide
    Call nucl1.SetProperties("21Ne")
    Call nucl2.SetProperties("10Be")
    N = 12
    Dim tickArr(3)
    tickArr(1) = 5
    tickArr(2) = 10
    tickArr(3) = 20
    '
    If logorlin = "lin" Then
        Call XYPlot(nucl1.name, nucl2.name, logorlin)
        EA1 = 0.000003
        EA2 = 0.00003
        EA3 = 0.005
        TA1 = 10000
        TA2 = 1500000
        TA3 = 25000000
        TB1 = 10000
        TB2 = 2000000
        TB3 = 5000000
        TB4 = 500000000
        TC1 = 10000
        TC2 = 2000000
        TC3 = 10000000
        TC4 = 500000000
        ED1 = 0.000005
        ED2 = 0.000015
        ED3 = 0.0001
        ED4 = 0.02
        Ee1 = 0.000005
        Ee2 = 0.000015
        EE3 = 0.0001
        EE4 = 0.02
        BA1 = 1000
        BA2 = 300000
        BA3 = 2000000
        BB1 = 1000
        BB2 = 1000000
        BB3 = 5000000
    Else
        tickFormula = getTickFormula(tickArr)
        Call XYPlot(nucl1.name, nucl2.name, tickFormula)
        Call addTickLabels(glob.xMin, tickArr)
        EA1 = 0.000003
        EA2 = 0.00003
        EA3 = 0.05
        TA1 = 1000
        TA2 = 1500000
        TA3 = 25000000
        TB1 = 1000
        TB2 = 500000
        TB3 = 1000000
        TB4 = 10000000
        TC1 = 1000
        TC2 = 2000000
        TC3 = 10000000
        TC4 = 100000000
        ED1 = 0.000003
        ED2 = 0.00003
        ED3 = 0.0002
        ED4 = 0.03
        Ee1 = 0.000005
        Ee2 = 0.00002
        EE3 = 0.0002
        EE4 = 0.03
        BA1 = 1
        BA2 = 500000
        BA3 = 5000000
        BB1 = 1
        BB2 = 500000
        BB3 = 5000000
    End If
    ' Plot saturation line
    Call PlotVariableErosion(N, "inf", 0, EA1, EA2, nucl1, nucl2, 1, 2, logorlin)
    Call PlotVariableErosion(N, "inf", 0, EA2, EA3, nucl1, nucl2, 1, 2, logorlin)
    '
    ' Plot zero erosion line
    Call PlotVariableTime(N, glob.Zero, 0, TA1, TA2, nucl1, nucl2, 1, 2, logorlin)
    Call PlotVariableTime(N, glob.Zero, 0, TA2, TA3, nucl1, nucl2, 1, 2, logorlin)
    '
    If glob.detail Then
        '
        If glob.zeroerosion Then
            '
            ' Plot 2Ma burial line
            Call PlotVariableTime(N, 0, 2000000, TB1, TB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0, 2000000, TB2, TB3, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0, 2000000, TB3, TB4, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 1Ma burial line
            Call PlotVariableTime(N, 0, 1000000, TB1, TB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0, 1000000, TB2, TB3, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0, 1000000, TB3, TB4, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 20ka exposure
            Call PlotVariableBurial(N, 20000, 0, BA1, BA2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, 20000, 0, BA2, BA3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 100ka exposure
            Call PlotVariableBurial(N, 100000, 0, BA1, BA2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, 100000, 0, BA2, BA3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 500ka exposure
            Call PlotVariableBurial(N, 500000, 0, BA1, BA2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, 500000, 0, BA2, BA3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after 2Ma exposure
            Call PlotVariableBurial(N, 2000000, 0, BA1, BA2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, 2000000, 0, BA2, BA3, nucl1, nucl2, 3, 1, logorlin)
            '
            Call getNaddLabel(20000, 0, 100000, "20ka", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(100000, 0, 100000, "100ka", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(500000, 0, 100000, "500ka", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(2000000, 0, 100000, "2Ma", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(50000, 0, 1000000, "1Ma burial", nucl1, nucl2, 10, logorlin)
            Call getNaddLabel(50000, 0, 2000000, "2Ma burial", nucl1, nucl2, 10, logorlin)
            '
        Else
            ' Plot 10m/Ma erosion line
            Call PlotVariableTime(N, 0.001, 0, TB1, TB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.001, 0, TB2, TB3, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.001, 0, TB3, TB4, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 2m/Ma erosion line
            Call PlotVariableTime(N, 0.0002, 0, TB1, TB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.0002, 0, TB2, TB3, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.0002, 0, TB3, TB4, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 0.5m/Ma erosion line
            Call PlotVariableTime(N, 0.00005, 0, TB1, TB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.00005, 0, TB2, TB3, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.00005, 0, TB3, TB4, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 0.1m/Ma erosion line
            Call PlotVariableTime(N, 0.00001, 0, TC1, TC2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.00001, 0, TC2, TC3, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableTime(N, 0.00001, 0, TC3, TC4, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot 500ka burial line
            Call PlotVariableErosion(N, "inf", 500000, ED1, ED2, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 500000, ED2, ED3, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 500000, ED3, ED4, nucl1, nucl2, 10, 1, logorlin)
            '
            ' Plot 1Ma burial line
            Call PlotVariableErosion(N, "inf", 1000000, Ee1, Ee2, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 1000000, Ee2, EE3, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 1000000, EE3, EE4, nucl1, nucl2, 10, 1, logorlin)
            '
            ' Plot 2Ma burial line
            Call PlotVariableErosion(N, "inf", 2000000, Ee1, Ee2, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 2000000, Ee2, EE3, nucl1, nucl2, 10, 1, logorlin)
            Call PlotVariableErosion(N, "inf", 2000000, EE3, EE4, nucl1, nucl2, 10, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 0.1m/Ma
            Call PlotVariableBurial(N, "inf", 0.00001, BA1, BA2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, "inf", 0.00001, BA2, BA3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 0.5m/Ma
            Call PlotVariableBurial(N, "inf", 0.00005, BB1, BB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, "inf", 0.00005, BB2, BB3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 2m/Ma
            Call PlotVariableBurial(N, "inf", 0.0002, BB1, BB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, "inf", 0.0002, BB2, BB3, nucl1, nucl2, 3, 1, logorlin)
            '
            ' Plot burial decay line after steady state erosion at 10m/Ma
            Call PlotVariableBurial(N, "inf", 0.001, BB1, BB2, nucl1, nucl2, 3, 1, logorlin)
            Call PlotVariableBurial(N, "inf", 0.001, BB2, BB3, nucl1, nucl2, 3, 1, logorlin)
            '
            Call getNaddLabel("inf", 0.00005, 750000, "0.05cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel("inf", 0.0002, 750000, "0.2cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel("inf", 0.001, 750000, "1cm/ka", nucl1, nucl2, 3, logorlin)
            Call getNaddLabel(5000000, glob.Zero, 0, "5Ma", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel(2000000, glob.Zero, 0, "2Ma", nucl1, nucl2, 7, logorlin)
            Call getNaddLabel("inf", 0.0005, 500000, "500ka burial", nucl1, nucl2, 10, logorlin)
            Call getNaddLabel("inf", 0.0005, 1000000, "1Ma burial", nucl1, nucl2, 10, logorlin)
            Call getNaddLabel("inf", 0.0005, 2000000, "2Ma burial", nucl1, nucl2, 10, logorlin)
            '
        End If
        '
    End If
    Set nucl1 = Nothing
    Set nucl2 = Nothing
End Sub
Private Sub PlotVariableErosion(ByVal N As Integer, ByVal time As Variant, ByVal burialT As Double, ByVal minE As Double, ByVal maxE As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByVal color As Integer, ByVal weight As Integer, ByVal logorlin As String)
    logMinE = Log(minE)
    logMaxE = Log(maxE)
    dLogE = (logMaxE - logMinE) / N
    formula = "=SERIES(,{"
    x = ""
    y = ""
    For i = 0 To N - 1
        E = Exp(logMinE + i * dLogE)
        N1 = getN(E, time, burialT, nuclide1)
        N2 = getN(E, time, burialT, nuclide2)
        x = x & Format(N2, "####E+0") & ","
        If logorlin = "lin" Then
            y = y & Format(N1 / N2, "####E+0") & ","
        ElseIf logorlin = "log" Then
            y = y & Format(Log(N1 / N2), "####E+0;-####E+0") & ","
        End If
    Next i
    E = maxE
    N1 = getN(E, time, burialT, nuclide1)
    N2 = getN(E, time, burialT, nuclide2)
    x = x & Format(N2, "####E+0")
    If logorlin = "lin" Then
        y = y & Format(N1 / N2, "####E+0")
    ElseIf logorlin = "log" Then
        y = y & Format(Log(N1 / N2), "####E+0;-####E+0")
    End If
    formula = formula & x & "},{" & y & "},1)"
    Call AddFormula(formula, ActiveChart.SeriesCollection.Count + 1, color, weight)
End Sub
Private Sub PlotVariableTime(ByVal N As Integer, ByVal E As Double, ByVal burialT As Variant, ByVal mint As Double, ByVal maxt As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByVal color As Integer, ByVal weight As Integer, ByVal logorlin As String)
    Dim time As Double
    
    logMinT = Log(mint)
    logMaxT = Log(maxt)
    dLogT = (logMaxT - logMinT) / N
    formula = "=SERIES(,{"
    x = ""
    y = ""
    For i = 0 To N - 1
        time = Exp(logMinT + i * dLogT)
        N1 = getN(E, time, burialT, nuclide1)
        N2 = getN(E, time, burialT, nuclide2)
        x = x & Format(N2, "####E+0") & ","
        If logorlin = "lin" Then
            y = y & Format(N1 / N2, "####E+0") & ","
        ElseIf logorlin = "log" Then
            y = y & Format(Log(N1 / N2), "####E+0;-####E+0") & ","
        End If
    Next i
    time = maxt
    N1 = getN(E, time, burialT, nuclide1)
    N2 = getN(E, time, burialT, nuclide2)
    x = x & Format(N2, "####E+0")
    If logorlin = "lin" Then
        y = y & Format(N1 / N2, "####E+0")
    ElseIf logorlin = "log" Then
        y = y & Format(Log(N1 / N2), "####E+0;-####E+0")
    End If
    formula = formula & x & "},{" & y & "},1)"
    Call AddFormula(formula, ActiveChart.SeriesCollection.Count + 1, color, weight)
End Sub
Private Sub PlotVariableBurial(ByVal N As Integer, ByVal time As Variant, ByVal E As Double, ByVal mint As Double, ByVal maxt As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByVal color As Integer, ByVal weight As Integer, ByVal logorlin As String)
    logMinT = Log(mint)
    logMaxT = Log(maxt)
    dLogT = (logMaxT - logMinT) / N
    formula = "=SERIES(,{"
    x = ""
    y = ""
    For i = 0 To N - 1
        burialT = Exp(logMinT + i * dLogT)
        N1 = getN(E, time, burialT, nuclide1)
        N2 = getN(E, time, burialT, nuclide2)
        x = x & Format(N2, "####E+0") & ","
        If logorlin = "lin" Then
            y = y & Format(N1 / N2, "####E+0") & ","
        ElseIf logorlin = "log" Then
            y = y & Format(Log(N1 / N2), "####E+0;-####E+0") & ","
        End If
    Next i
    burialT = maxt
    N1 = getN(E, time, burialT, nuclide1)
    N2 = getN(E, time, burialT, nuclide2)
    x = x & Format(N2, "####E+0")
    If logorlin = "lin" Then
        y = y & Format(N1 / N2, "####E+0")
    ElseIf logorlin = "log" Then
        y = y & Format(Log(N1 / N2), "####E+0;-####E+0")
    End If
    formula = formula & x & "},{" & y & "},1)"
    Call AddFormula(formula, ActiveChart.SeriesCollection.Count + 1, color, weight)
End Sub
Private Sub AddLabel(ByVal N1 As Double, ByVal N2 As Double, ByVal Label As String, ByVal color As Integer, ByVal placement As String, ByVal logorlin As String)
    If logorlin = "lin" Then
        formula = "=SERIES(,{" & Format(N2, "####E+0") & "},{" & Format(N1 / N2, "####E+0") & "},1)"
    Else
        formula = "=SERIES(,{" & Format(N2, "####E+0") & "},{" & Format(Log(N1 / N2), "####E+0;-####E+0") & "},1)"
    End If
    newseriesnum = ActiveChart.SeriesCollection.Count + 1
    Call AddFormula(formula, newseriesnum, 1, 1)
    ActiveChart.SeriesCollection(1).name = Label
    If Label <> "" Then
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).Points(1).DataLabel.Text = Label
        If placement = "center" Then
            ActiveChart.SeriesCollection(Label).DataLabels.Position = xlLabelPositionCenter
        ElseIf placement = "left" Then
            ActiveChart.SeriesCollection(Label).DataLabels.Position = xlLabelPositionLeft
        End If
        ActiveChart.SeriesCollection(Label).DataLabels.Font.ColorIndex = color
    End If
End Sub
Private Sub getNaddLabel(ByVal time As Variant, ByVal E As Double, ByVal burialT As Double, ByVal Label As String, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByVal color As Integer, ByVal logorlin As String)
    N1 = getN(E, time, burialT, nuclide1)
    N2 = getN(E, time, burialT, nuclide2)
    Call AddLabel(N1, N2, Label, color, "center", logorlin)
End Sub
Private Sub addTickLabels(ByVal x As Double, ByVal Arr As Variant)
    For i = 1 To UBound(Arr)
        Call AddLabel(Arr(i) * x, x, Arr(i), 1, "left", "log")
    Next i
End Sub
Private Sub XYPlot(ByVal N1label As String, ByVal N2label As String, ByVal linOrLabel As Variant)
    Dim mySrs As Series
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatter
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.Location Where:=xlLocationAsNewSheet
    With ActiveChart.Axes(xlCategory)
        .MaximumScaleIsAuto = True
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .Crosses = xlAutomatic
        .ReversePlotOrder = False
        '.ScaleType = xlLogarithmic
        .DisplayUnit = xlNone
    End With
    With ActiveChart.Axes(xlValue)
        .MinorUnitIsAuto = True
        .MajorUnitIsAuto = True
        .ReversePlotOrder = False
        '.ScaleType = xlLinear
        .DisplayUnit = xlNone
    End With
    '
    If linOrLabel = "lin" Then
        ActiveChart.Axes(xlCategory).ScaleType = xlLinear
    Else
        ActiveChart.Axes(xlValue).MajorTickMark = xlNone
        ActiveChart.Axes(xlValue).MinorTickMark = xlNone
        ActiveChart.Axes(xlValue).TickLabelPosition = xlNone
        ActiveChart.Axes(xlCategory).ScaleType = xlLogarithmic
    End If
    '
    ActiveChart.Axes(xlCategory).MinimumScaleIsAuto = False
    ActiveChart.Axes(xlCategory).MinimumScale = glob.xMin
    '
    ActiveChart.Axes(xlCategory).MaximumScaleIsAuto = False
    ActiveChart.Axes(xlCategory).MaximumScale = glob.xMax
    '
    ActiveChart.Axes(xlValue).MinimumScaleIsAuto = False
    ActiveChart.Axes(xlValue).Crosses = xlCustom
    If linOrLabel = "lin" Then
        ActiveChart.Axes(xlValue).MinimumScale = glob.yMin
        ActiveChart.Axes(xlValue).CrossesAt = glob.yMin
    Else
        ActiveChart.Axes(xlValue).MinimumScale = Log(glob.yMin)
        ActiveChart.Axes(xlValue).CrossesAt = Log(glob.yMin)
    End If
    '
    ActiveChart.Axes(xlValue).MaximumScaleIsAuto = False
    If linOrLabel = "lin" Then
        ActiveChart.Axes(xlValue).MaximumScale = glob.yMax
    Else
        ActiveChart.Axes(xlValue).MaximumScale = Log(glob.yMax)
    End If
    '
    ActiveChart.Axes(xlValue).MajorGridlines.Delete
    ActiveChart.PlotArea.ClearFormats
    ActiveChart.Legend.Delete
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = N2label
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = N1label & " / " & N2label
    End With
    '
    With ActiveChart.PlotArea.Border
        .weight = xlHairline
        .LineStyle = xlContinuous
    End With
    With ActiveChart.PlotArea.Interior
        .color = RGB(255, 255, 255)
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With
    With ActiveChart.ChartArea.Border
        .weight = xlHairline
        .LineStyle = xlNone
    End With
    ActiveChart.ChartArea.Shadow = False
    With ActiveChart.ChartArea.Interior
        .color = RGB(200, 200, 200) 'RGB(255, 153, 204)
        .PatternColorIndex = 1
        .Pattern = xlSolid
    End With
    '
    For Each mySrs In ActiveChart.SeriesCollection
        mySrs.Delete
    Next
    If linOrLabel <> "lin" Then
        Call addTicks(linOrLabel)
    End If
End Sub
Private Sub addTicks(ByVal tickFormula As String)
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count).formula = tickFormula
    With ActiveChart.SeriesCollection(1).Border
        .ColorIndex = 1
        .weight = xlThin
        .LineStyle = xlNone
    End With
    With ActiveChart.SeriesCollection(1)
        .MarkerBackgroundColorIndex = xlNone
        .MarkerForegroundColorIndex = xlNone
        .MarkerStyle = xlNone
        .Smooth = False
        .MarkerSize = 5
        .Shadow = False
    End With
    ActiveChart.SeriesCollection(1).ErrorBar Direction:=xlX, Include:= _
        xlPlusValues, Type:=xlFixedValue, Amount:=glob.xMin / 10
    With ActiveChart.SeriesCollection(1).ErrorBars.Border
        .LineStyle = xlContinuous
        .ColorIndex = 57
        .weight = xlThin
    End With
    ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
End Sub
Private Sub AddFormula(ByVal formula As String, ByVal seriesnum As Integer, ByVal color As Integer, ByVal weight As Integer)
    If ActiveChart.SeriesCollection.Count < seriesnum Then
        ActiveChart.SeriesCollection.NewSeries
    End If
    ActiveChart.SeriesCollection(seriesnum).formula = formula
    With ActiveChart.SeriesCollection(1).Border
        .ColorIndex = color
        If weight = 1 Then
            .weight = xlThin
        ElseIf weight = 2 Then
            .weight = xlMedium
        Else
            .weight = xlThick
        End If
        .LineStyle = xlContinuous
    End With
    With ActiveChart.SeriesCollection(1)
        .MarkerBackgroundColorIndex = xlAutomatic
        .MarkerForegroundColorIndex = xlAutomatic
        .MarkerStyle = xlNone
        .Smooth = True
        .MarkerSize = 5
        .Shadow = False
    End With
End Sub
Private Sub DataPlot(ByVal theSheet As String, theRange As MyRange, ByVal logorlin As String)
    Dim CurrSeries As Integer
    Dim row As Range
    Dim Yerr As Double, Seff1 As Double, Seff2 As Double
    If (theRange.numcols = 6) Then
        firstcol = 0
    ElseIf (theRange.numcols = 7) Then
        firstcol = 1
    End If
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        N1 = theRange.CellValue(rownum, firstcol + 2)
        If N1 <> "" And IsNumeric(N1) Then
            N2 = theRange.CellValue(rownum, firstcol + 5)
            S1 = theRange.CellValue(rownum, firstcol + 1)
            S2 = theRange.CellValue(rownum, firstcol + 4)
            Call getSeff(S1, S2, N1, N2, Seff1, Seff2)
            N1 = theRange.CellValue(rownum, firstcol + 2) / Seff1
            N2 = theRange.CellValue(rownum, firstcol + 5) / Seff2
            N1err = theRange.CellValue(rownum, firstcol + 3) / Seff1
            N2err = theRange.CellValue(rownum, firstcol + 6) / Seff2
            If glob.PlotEllipse Then
                Call ErrorEllipse(N1, N2, N1err, N2err, logorlin)
            Else
                Call PlotPoint(N1, N2, logorlin)
                Call ErrorBars(N1, N2, N1err, N2err, logorlin)
            End If
            If (theRange.numcols = 7) Then
                Call AddLabel(N1, N2, theRange.CellValue(rownum, firstcol), 1, "center", logorlin)
            End If
        End If
    Next rownum
End Sub
Private Sub getSeff(ByVal S1 As Double, ByVal S2 As Double, ByVal N1 As Double, ByVal N2 As Double, ByRef Seff1 As Double, ByRef Seff2 As Double)

    If (glob.zeroerosion) Then
        Seff1 = S1
        Seff2 = S2
    Else

    Dim nucl1 As MyNuclide, nucl2 As MyNuclide, Eae As Double, Ebe As Double, Tae As Double, Tbe As Double, Eerr As Double, terr As Double

    Set nucl1 = New MyNuclide
    Set nucl2 = New MyNuclide
    If BananaForm.AlBeOptionButton.Value Then
        Call nucl1.SetProperties("26Al")
    ElseIf BananaForm.NeBeOptionButton.Value Then
        Call nucl1.SetProperties("21Ne")
    End If
    Call nucl2.SetProperties("10Be")
    Call nucl1.SetScaling(S1)
    Call nucl2.SetScaling(S2)
    '
    Call getAgeErosion(N1, N1 / 10, N2, N2 / 10, nucl1, nucl2, Eae, Eerr, Tae, terr)
    N1ae = getN(Eae, Abs(Tae), 0, nucl1)
    N2ae = getN(Eae, Abs(Tae), 0, nucl2)
    aeMisfit = Abs((N1ae - N1) / N1) + Abs((N2ae - N2) / N2)
    '
    Call getBurialErosion(N1, N1 / 10, N2, N2 / 10, nucl1, nucl2, Ebe, Eerr, Tbe, terr)
    N1be = getN(Ebe, "inf", Abs(Tbe), nucl1)
    N2be = getN(Ebe, "inf", Abs(Tbe), nucl2)
    beMisfit = Abs((N1be - N1) / N1) + Abs((N2be - N2) / N2)
    '
    Ee1 = getErosion(N1, nucl1)
    Ee2 = getErosion(N2, nucl2)
    N1e = getN(Ee1, "inf", 0, nucl1)
    N2e = getN(Ee2, "inf", 0, nucl2)
    '
    Call nucl1.SetScaling(1)
    Call nucl2.SetScaling(1)
    If ((Tae > 0) And (Eae > 0) And (aeMisfit < 0.1)) Then
        N1pred = getN(Eae, Tae, 0, nucl1)
        N2pred = getN(Eae, Tae, 0, nucl2)
    ElseIf ((Tbe > 0) And (Ebe > 0) And (beMisfit < 0.1)) Then
        N1pred = getN(Ebe, "inf", Tbe, nucl1)
        N2pred = getN(Ebe, "inf", Tbe, nucl2)
    Else
        N1pred = getN(Ee1, "inf", 0, nucl1)
        N2pred = getN(Ee2, "inf", 0, nucl2)
    End If
    Seff1 = N1 / N1pred
    Seff2 = N2 / N2pred
    Set nucl1 = Nothing
    Set nucl2 = Nothing
    
    End If
    
End Sub
Private Function getTickFormula(ByVal Arr As Variant)
    x = ""
    y = ""
    formula = "=SERIES(,{"
    For i = 1 To UBound(Arr) - 1
        x = x & Format(glob.xMin, "####E+0") & ","
        y = y & Format(Log(Arr(i)), "####E+0;-####E+0") & ","
    Next i
    x = x & Format(glob.xMin, "####E+0")
    y = y & Format(Log(Arr(i)), "####E+0;-####E+0")
    getTickFormula = formula & x & "},{" & y & "},1)"
End Function
Private Sub PlotPoint(N1, N2, logorlin)
    ActiveChart.SeriesCollection.NewSeries
    CurrSeries = ActiveChart.SeriesCollection.Count
    ActiveChart.SeriesCollection(CurrSeries).XValues = N2
    If logorlin = "lin" Then
        ActiveChart.SeriesCollection(CurrSeries).Values = N1 / N2
    Else
        ActiveChart.SeriesCollection(CurrSeries).Values = Log(N1 / N2)
    End If
    With ActiveChart.SeriesCollection(CurrSeries).Border
        .weight = xlHairline
        .LineStyle = xlNone
    End With
    With ActiveChart.SeriesCollection(CurrSeries)
        .MarkerBackgroundColorIndex = 2
        .MarkerForegroundColorIndex = 1
        .MarkerStyle = xlCircle
        .Smooth = False
        .MarkerSize = 8
        .Shadow = False
    End With
End Sub
Private Sub ErrorBars(ByVal N1 As Double, ByVal N2 As Double, ByVal N1err As Double, ByVal N2err As Double, ByVal logorlin As String)
    Dim CurrSeries As Integer

    x = N2
    Xerr = N2err
    y = N1 / N2
    Yerr = y * ((N1err / N1) ^ 2 + (N2err / N2) ^ 2) ^ 0.5
    CurrSeries = ActiveChart.SeriesCollection.Count
    If logorlin = "lin" Then
        ActiveChart.SeriesCollection(CurrSeries).ErrorBar Direction:=xlX, Include:=xlBoth, _
            Type:=xlCustom, Amount:=glob.sigma * Xerr, MinusValues:=glob.sigma * Xerr
    ElseIf logorlin = "log" Then
        Xerr = Xerr / x
        Yerr = Yerr / y
        x = Log(x)
        y = Log(y)
        ActiveChart.SeriesCollection(CurrSeries).ErrorBar Direction:=xlX, Include:=xlBoth, _
            Type:=xlCustom, Amount:=Exp(x + glob.sigma * Xerr) - Exp(x), MinusValues:=Exp(x) - Exp(x - glob.sigma * Xerr)
    End If
    ActiveChart.SeriesCollection(CurrSeries).ErrorBar Direction:=xlY, Include:=xlBoth, _
        Type:=xlCustom, Amount:=glob.sigma * Yerr, MinusValues:=glob.sigma * Yerr
End Sub
Private Sub ErrorEllipse(ByVal N1bar As Double, ByVal N2bar As Double, ByVal sigmaN1 As Double, ByVal sigmaN2 As Double, ByVal logorlin As String)
    Dim CurrSeries As Integer
    Dim rho, alpha, a, B As Double
    
    rho = -N1bar * sigmaN2 / (N2bar ^ 2 * sigmaN1 ^ 2 + N1bar ^ 2 * sigmaN2 ^ 2) ^ 0.5
    xbar = N2bar
    ybar = N1bar / N2bar
    sigmaX = sigmaN2
    sigmaY = ybar * ((sigmaN1 / N1bar) ^ 2 + (sigmaN2 / N2bar) ^ 2) ^ 0.5
    If logorlin = "log" Then
        sigmaX = sigmaX / xbar
        sigmaY = sigmaY / ybar
        xbar = Log(xbar)
        ybar = Log(ybar)
    End If
    alpha = 0.5 * Atn(2 * rho * sigmaX * sigmaY / (sigmaX ^ 2 - sigmaY ^ 2))
    a = (sigmaX ^ 2 * sigmaY ^ 2 * (1 - rho ^ 2) / (sigmaY ^ 2 * Cos(alpha) ^ 2 - 2 * rho * sigmaX * sigmaY * Sin(alpha) * Cos(alpha) + sigmaX ^ 2 * Sin(alpha) ^ 2)) ^ 0.5
    B = (sigmaX ^ 2 * sigmaY ^ 2 * (1 - rho ^ 2) / (sigmaY ^ 2 * Sin(alpha) ^ 2 + 2 * rho * sigmaX * sigmaY * Sin(alpha) * Cos(alpha) + sigmaX ^ 2 * Cos(alpha) ^ 2)) ^ 0.5
    Call plotEllipseSection(0, 2 * pi, glob.sigma * a, glob.sigma * B, alpha, xbar, ybar, logorlin)
End Sub
Private Sub plotEllipseSection(mint, maxt, a, B, alpha, xbar, ybar, ByVal logorlin As String)
    Dim color As Integer
    
    N = 14
    formula = "=SERIES(,{"
    x = ""
    y = ""
    For t = mint To maxt Step (maxt - mint) / N
        If logorlin = "lin" Then
            x = x & Format(a * Cos(t) * Cos(alpha) - B * Sin(t) * Sin(alpha) + xbar, "####E+0") & ","
        ElseIf logorlin = "log" Then
            x = x & Format(Exp(a * Cos(t) * Cos(alpha) - B * Sin(t) * Sin(alpha) + xbar), "####E+0;-####E+0") & ","
        End If
        y = y & Format(a * Cos(t) * Sin(alpha) + B * Sin(t) * Cos(alpha) + ybar, "####E+0") & ","
    Next t
    If logorlin = "lin" Then
        x = x & Format(a * Cos(maxt) * Cos(alpha) - B * Sin(maxt) * Sin(alpha) + xbar, "####E+0")
    ElseIf logorlin = "log" Then
        x = x & Format(Exp(a * Cos(maxt) * Cos(alpha) - B * Sin(maxt) * Sin(alpha) + xbar), "####E+0;-####E+0")
    End If
    y = y & Format(a * Cos(maxt) * Sin(alpha) + B * Sin(maxt) * Cos(alpha) + ybar, "####E+0")
    formula = formula & x & "},{" & y & "},1)"
    Call AddFormula(formula, ActiveChart.SeriesCollection.Count + 1, 1, 1)
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub HelpButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Select a range of cells, six or seven columns wide " & vbCrLf & vbCrLf
    Msg = Msg & "    Label(1) Sc1(1) N1(1) Nerr1(1) Sc2(1) N2(1) Nerr2(1)" & vbCrLf
    Msg = Msg & "       :      :      :      :       :      :              " & vbCrLf
    Msg = Msg & "    Label(n) Sc1(n) N1(n) Nerr1(n) Sc2(n) N2(n) Nerr2(n)" & vbCrLf & vbCrLf
    Msg = Msg & "with:         " & vbCrLf
    Msg = Msg & "  - Label = Sample name (optional)                   " & vbCrLf
    Msg = Msg & "  - Sc1 = Composite scaling factor for Nuclide 1, " & vbCrLf
    Msg = Msg & "          i.e. the combined effects of latitude," & vbCrLf
    Msg = Msg & "          elevation, snow- and self-shielding" & vbCrLf
    Msg = Msg & "  - N1 = Concentration of Nuclide 1 (atoms/gram)" & vbCrLf
    Msg = Msg & "  - Nerr1 = 1-sigma uncertainty for N1 (atoms/gram) " & vbCrLf
    Msg = Msg & "  - Sc2 = Composite scaling factor for Nuclide 2" & vbCrLf
    Msg = Msg & "  - N2 = concentration of Nuclide 2 (atoms/gram)" & vbCrLf
    Msg = Msg & "  - Nerr2 = 1-sigma uncertainty for N2 (atoms/gram) " & vbCrLf & vbCrLf
    Msg = Msg & "N.B: N1 and N2 must be topography-corrected, unless" & vbCrLf
    Msg = Msg & "     the topographic shielding correction is small (S_t > 0.95)" & vbCrLf
    If Not glob.PC Then
        Msg = Msg & "To add more detail to the banana plots, see Options" & vbCrLf
    End If
    MsgBox Msg, vbInformation, APPNAME
End Sub
