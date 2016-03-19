VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyGlobals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_name As String
Private m_version As String
Private m_rho As Double
Private m_PC As Boolean
Private m_detail As Boolean
Private m_zeroerosion As Boolean
Private m_L3He As Double
Private m_L10Be As Double
Private m_L14C As Double
Private m_L21Ne As Double
Private m_L26Al As Double
Private m_L36Cl As Double
Private m_AlBeLogOrLin As String
Private m_NeBeLogOrLin As String
Private m_Replace As Boolean
Private m_PlotEllipse As Boolean
Private m_NewtonOption As Boolean
Private m_MetropIter As Integer
Private m_ConfiLevel As Double
Private m_T_o As Double
Private m_B_o As Double
Private m_R_d As Double
Private m_G_o As Double
Private m_P_o As Double
Private m_MM0 As Double
Private m_Zero As Double
Private m_L0 As Double
Private m_L1 As Double
Private m_L2 As Double
Private m_L3 As Double
Private m_P3He As Double
Private m_F3He0 As Double
Private m_F3He1 As Double
Private m_F3He2 As Double
Private m_F3He3 As Double
Private m_P21Ne As Double
Private m_F21Ne0 As Double
Private m_F21Ne1 As Double
Private m_F21Ne2 As Double
Private m_F21Ne3 As Double
Private m_P14C As Double
Private m_F14C0 As Double
Private m_F14C1 As Double
Private m_F14C2 As Double
Private m_F14C3 As Double
Private m_P26Al As Double
Private m_F26Al0 As Double
Private m_F26Al1 As Double
Private m_F26Al2 As Double
Private m_F26Al3 As Double
Private m_P10Be As Double
Private m_F10Be0 As Double
Private m_F10Be1 As Double
Private m_F10Be2 As Double
Private m_F10Be3 As Double
Private m_P36Cl As Double
Private m_F36Cl0 As Double
Private m_F36Cl1 As Double
Private m_F36Cl2 As Double
Private m_F36Cl3 As Double
Private m_P21Ne10Be As Double
Private m_scaling As String
Private m_equation As String
Private m_xMin As Double
Private m_xMax As Double
Private m_yMin As Double
Private m_yMax As Double
Private m_exponent As Double
Private m_sigma As Double
Private m_n10BeCals As Integer
Private m_n26AlCals As Integer
Private m_n36ClCals As Integer
Private m_n14CCals As Integer
Private m_n3HeCals As Integer
Private m_n21NeCals As Integer
Private m_tieNe2Be As Boolean
Private Sub class_initialize()
    On Error Resume Next
    Dim currSheet As String, versie As String
    currSheet = ActiveSheet.name
    m_name = "CosmoCalcVars"
    m_PC = True
    versie = "version 2.4"
    Dim i As Integer, i2 As Integer
    Dim bWksExists As Boolean
    i2 = Worksheets.Count
    For i = 1 To i2
    If Worksheets(i).name = m_name Then
        bWksExists = True
    End If
    Next i
    If (Not bWksExists) Then
        Worksheets.Add After:=Worksheets(i2)
        Worksheets(i2 + 1).name = m_name
    End If
    If (Not bWksExists) Then
        Call createNew(versie)
    ElseIf (Worksheets(m_name).Range("version").Value <> versie) Then
        Call createNew(versie)
    Else
        With Worksheets(m_name)
            m_version = .Range("version").Value
            m_detail = .Range("detail").Value
            m_zeroerosion = .Range("zeroerosion").Value
            m_L3He = .Range("L_3He").Value
            m_L10Be = .Range("L_10Be").Value
            m_L14C = .Range("L_14C").Value
            m_L21Ne = .Range("L_21Ne").Value
            m_L26Al = .Range("L_26Al").Value
            m_L36Cl = .Range("L_36Cl").Value
            m_AlBeLogOrLin = .Range("AlBeLogOrLin").Value
            m_NeBeLogOrLin = .Range("NeBeLogOrLin").Value
            m_Replace = .Range("Replace").Value
            m_PlotEllipse = .Range("PlotEllipse").Value
            m_NewtonOption = .Range("NewtonOption").Value
            m_MetropIter = .Range("MetropIter").Value
            m_ConfiLevel = .Range("ConfiLevel").Value
            m_T_o = .Range("T_o").Value
            m_B_o = .Range("B_o").Value
            m_R_d = .Range("R_d").Value
            m_G_o = .Range("G_o").Value
            m_P_o = .Range("P_o").Value
            m_MM0 = .Range("MM_0").Value
            m_Zero = .Range("Zero").Value
            m_L0 = .Range("L_0").Value
            m_L1 = .Range("L_1").Value
            m_L2 = .Range("L_2").Value
            m_L3 = .Range("L_3").Value
            m_F3He0 = .Range("F_3He0").Value
            m_F3He1 = .Range("F_3He1").Value
            m_F3He2 = .Range("F_3He2").Value
            m_F3He3 = .Range("F_3He3").Value
            m_F21Ne0 = .Range("F_21Ne0").Value
            m_F21Ne1 = .Range("F_21Ne1").Value
            m_F21Ne2 = .Range("F_21Ne2").Value
            m_F21Ne3 = .Range("F_21Ne3").Value
            m_F14C0 = .Range("F_14C0").Value
            m_F14C1 = .Range("F_14C1").Value
            m_F14C2 = .Range("F_14C2").Value
            m_F14C3 = .Range("F_14C3").Value
            m_F26Al0 = .Range("F_26Al0").Value
            m_F26Al1 = .Range("F_26Al1").Value
            m_F26Al2 = .Range("F_26Al2").Value
            m_F26Al3 = .Range("F_26Al3").Value
            m_F10Be0 = .Range("F_10Be0").Value
            m_F10Be1 = .Range("F_10Be1").Value
            m_F10Be2 = .Range("F_10Be2").Value
            m_F10Be3 = .Range("F_10Be3").Value
            m_F36Cl0 = .Range("F_36Cl0").Value
            m_F36Cl1 = .Range("F_36Cl1").Value
            m_F36Cl2 = .Range("F_36Cl2").Value
            m_F36Cl3 = .Range("F_36Cl3").Value
            m_P3He = .Range("P_3He").Value
            m_P21Ne = .Range("P_21Ne").Value
            m_P14C = .Range("P_14C").Value
            m_P26Al = .Range("P_26Al").Value
            m_P10Be = .Range("P_10Be").Value
            m_P36Cl = .Range("P_36Cl").Value
            m_P21Ne10Be = .Range("P_21Ne10Be").Value
            m_scaling = .Range("scaling").Value
            m_equation = .Range("equation").Value
            m_xMin = .Range("xMin").Value
            m_xMax = .Range("xMax").Value
            m_yMin = .Range("yMin").Value
            m_yMax = .Range("yMax").Value
            m_exponent = .Range("exponent").Value
            m_rho = .Range("rho").Value
            m_sigma = .Range("sigma").Value
            m_n10BeCals = .Range("n_10BeCals").Value
            m_n26AlCals = .Range("n_26AlCals").Value
            m_n3HeCals = .Range("n_3HeCals").Value
            m_n14CCals = .Range("n_14CCals").Value
            m_n36ClCals = .Range("n_36ClCals").Value
            m_n21NeCals = .Range("n_21NeCals").Value
            m_tieNe2Be = .Range("tieNe2Be").Value
        End With
    End If
    Worksheets(currSheet).Activate
End Sub
Private Sub createNew(ByVal versie As String)
        With Worksheets(m_name)
            .Visible = xlHidden
            .Range("A1").name = "version"
            .Range("A2").name = "rho"
            .Range("A3").name = "L_3He"
            .Range("A4").name = "L_10Be"
            .Range("A5").name = "L_14C"
            .Range("A6").name = "L_21Ne"
            .Range("A7").name = "L_26Al"
            .Range("A8").name = "L_36Cl"
            .Range("A9").name = "L_0"
            .Range("A10").name = "L_1"
            .Range("A11").name = "L_2"
            .Range("A12").name = "L_3"
            .Range("A13").name = "F_3He0"
            .Range("A14").name = "F_3He1"
            .Range("A15").name = "F_3He2"
            .Range("A16").name = "F_3He3"
            .Range("A17").name = "F_21Ne0"
            .Range("A18").name = "F_21Ne1"
            .Range("A19").name = "F_21Ne2"
            .Range("A20").name = "F_21Ne3"
            .Range("A21").name = "F_14C0"
            .Range("A22").name = "F_14C1"
            .Range("A23").name = "F_14C2"
            .Range("A24").name = "F_14C3"
            .Range("A25").name = "F_26Al0"
            .Range("A26").name = "F_26Al1"
            .Range("A27").name = "F_26Al2"
            .Range("A28").name = "F_26Al3"
            .Range("A29").name = "F_10Be0"
            .Range("A30").name = "F_10Be1"
            .Range("A31").name = "F_10Be2"
            .Range("A32").name = "F_10Be3"
            .Range("A33").name = "F_36Cl0"
            .Range("A34").name = "F_36Cl1"
            .Range("A35").name = "F_36Cl2"
            .Range("A36").name = "F_36Cl3"
            .Range("A37").name = "P_3He"
            .Range("A38").name = "P_21Ne"
            .Range("A39").name = "P_14C"
            .Range("A40").name = "P_26Al"
            .Range("A41").name = "P_10Be"
            .Range("A42").name = "P_36Cl"
            .Range("A43").name = "P_21Ne10Be"
            .Range("A44").name = "scaling"
            .Range("A45").name = "equation"
            .Range("A46").name = "T_o"
            .Range("A47").name = "B_o"
            .Range("A48").name = "R_d"
            .Range("A49").name = "G_o"
            .Range("A50").name = "P_o"
            .Range("A51").name = "MM_0"
            .Range("A52").name = "exponent"
            .Range("A53").name = "sigma"
            .Range("A54").name = "Zero"
            .Range("A55").name = "xMin"
            .Range("A56").name = "xMax"
            .Range("A57").name = "yMin"
            .Range("A58").name = "yMax"
            .Range("A59").name = "AlBeLogOrLin"
            .Range("A60").name = "NeBeLogOrLin"
            .Range("A61").name = "Replace"
            .Range("A62").name = "detail"
            .Range("A63").name = "zeroerosion"
            .Range("A64").name = "PlotEllipse"
            .Range("A65").name = "NewtonOption"
            .Range("A66").name = "MetropIter"
            .Range("A67").name = "ConfiLevel"
            .Range("A68").name = "n_10BeCals"
            .Range("A69").name = "n_26AlCals"
            .Range("A70").name = "n_21NeCals"
            .Range("A71").name = "n_3HeCals"
            .Range("A72").name = "n_14CCals"
            .Range("A73").name = "n_36ClCals"
            .Range("A74").name = "tieNe2Be"
        End With
        Me.version = versie
        Call reset
End Sub
Public Sub setSLHLp(nuclName As String)
    Dim S As Double, Lat As Double, Elev As Double, N As Double, Age As Double
    Dim nucl As MyNuclide
    Set nucl = New MyNuclide
    Dim rng As MyRange
    Set rng = New MyRange
    On Error Resume Next
    Call rng.SetProperties(Me.LatElevRange(nuclName), Me.name)
    Call nucl.SetProperties(nuclName)
    Select Case Me.Scaling
    Case Is = "Lal"
        Call LalScaling(rng)
    Case Is = "Stone"
        Call StoneScaling(rng, nucl)
    Case Is = "Dunai"
        Call DunaiScaling(rng, nucl)
    Case Is = "Desilets & Zreda (2003)"
        Call DesiletScaling(2003, rng, nucl)
    Case Is = "Desilets et al (2006)"
        Call DesiletScaling(2006, rng, nucl)
    End Select
    Call rng.SetProperties(Me.CalRange(nuclName), Me.name)
    With Worksheets(Me.name)
    mySum = 0
    For rownum = 1 To rng.numRows
        N = .Range(rng.CellAddress(rownum, 1)).Value
        Age = .Range(rng.CellAddress(rownum, 2)).Value
        S = .Range(rng.CellAddress(rownum, 5)).Value
        P = getP(N, Age, S, nucl)
        .Range(rng.CellAddress(rownum, 6)).Value = P
        mySum = mySum + P
    Next rownum
    Select Case nuclName
    Case Is = "10Be"
        Me.P10Be = mySum / rng.numRows
    Case Is = "26Al"
        Me.P26Al = mySum / rng.numRows
    Case Is = "3He"
        Me.P3He = mySum / rng.numRows
    Case Is = "14C"
        Me.P14C = mySum / rng.numRows
    Case Is = "21Ne"
        If glob.tieNe2Be Then
            Me.P21Ne = Me.P10Be * Me.P21Ne10Be
        Else
            Me.P21Ne = mySum / rng.numRows
        End If
    Case Is = "36Cl"
        Me.P36Cl = mySum / rng.numRows
    End Select
    End With
    Set rng = Nothing
    Set nucl = Nothing
End Sub
Public Sub ConvertAllP(oldScaling As String, newScaling As String)
    Call ConvertP(oldScaling, newScaling, "10Be")
    Call ConvertP(oldScaling, newScaling, "26Al")
    Call ConvertP(oldScaling, newScaling, "21Ne")
    Call ConvertP(oldScaling, newScaling, "3He")
    Call ConvertP(oldScaling, newScaling, "14C")
    Call ConvertP(oldScaling, newScaling, "36Cl")
End Sub
Private Sub ConvertP(oldScaling As String, newScaling As String, nucl As String)
    Dim didReplace As Boolean
    Dim latRng As MyRange
    Set latRng = New MyRange
    Dim elevRng As MyRange
    Set elevRng = New MyRange
    Call latRng.SetProperties(Me.latRange(nucl), Me.name)
    Call elevRng.SetProperties(Me.elevRange(nucl), Me.name)
    didReplace = Me.Replace
    Me.Replace = True
    Select Case oldScaling
        Case Is = "Lal"
            Select Case newScaling
                Case Is = "Lal"
                    ' do nothing
                Case Is = "Stone"
                    Call zToP(elevRng)
                Case Is = "Dunai"
                    Call degToI(latRng)
                    Call zToD(elevRng)
                Case Is = "Desilets & Zreda (2003)"
                    Call degToRc(latRng)
                    Call zToD(elevRng)
                Case Is = "Desilets et al (2006)"
                    Call degToRc(latRng)
                    Call zToD(elevRng)
            End Select
        Case Is = "Stone"
            Select Case newScaling
                Case Is = "Lal"
                    Call pToZ(elevRng)
                Case Is = "Stone"
                    ' do nothing
                Case Is = "Dunai"
                    Call degToI(latRng)
                    Call pToD(elevRng)
                Case Is = "Desilets & Zreda (2003)"
                    Call degToRc(latRng)
                    Call pToD(elevRng)
                Case Is = "Desilets et al (2006)"
                    Call degToRc(latRng)
                    Call pToD(elevRng)
            End Select
        Case Is = "Dunai"
            Select Case newScaling
                Case Is = "Lal"
                    Call iToDeg(latRng)
                    Call dToZ(elevRng)
                Case Is = "Stone"
                    Call iToDeg(latRng)
                    Call dToP(elevRng)
                Case Is = "Dunai"
                    ' do nothing
                Case Is = "Desilets & Zreda (2003)"
                    Call iToDeg(latRng)
                    Call degToRc(latRng)
                Case Is = "Desilets et al (2006)"
                    Call iToDeg(latRng)
                    Call degToRc(latRng)
            End Select
        Case Is = "Desilets & Zreda (2003)"
            Select Case newScaling
                Case Is = "Lal"
                    Call RcToDeg(latRng)
                    Call dToZ(elevRng)
                Case Is = "Stone"
                    Call RcToDeg(latRng)
                    Call dToP(elevRng)
                Case Is = "Dunai"
                    Call RcToDeg(latRng)
                    Call degToI(latRng)
                Case Is = "Desilets & Zreda (2003)"
                    'do nothing
                Case Is = "Desilets et al (2006)"
                    'do nothing
            End Select
        Case Is = "Desilets et al (2006)"
            Select Case newScaling
                Case Is = "Lal"
                    Call RcToDeg(latRng)
                    Call dToZ(elevRng)
                Case Is = "Stone"
                    Call RcToDeg(latRng)
                    Call dToP(elevRng)
                Case Is = "Dunai"
                    Call RcToDeg(latRng)
                    Call degToI(latRng)
                Case Is = "Desilets & Zreda (2003)"
                    'do nothing
                Case Is = "Desilets et al (2006)"
                    'do nothing
            End Select
    End Select
    Me.Scaling = newScaling
    Call Me.setSLHLp(nucl)
    Me.Replace = didReplace
    Set latRng = Nothing
    Set elevRng = Nothing
End Sub
Private Function getP(N As Double, t As Double, S As Double, nucl As MyNuclide) As Double
    If nucl.L = 0 Then
        getP = N / (S * t)
    Else
        getP = N * nucl.L / (S * (1 - Exp(-nucl.L * t)))
    End If
End Function
Public Sub setSLHLf(ByVal equation As String)
    With Me
        .L0 = 160
        .L1 = 738.6
        .L2 = 2688
        .L3 = 4360
        .F3He0 = 1
        .F3He1 = 0
        .F3He2 = 0
        .F3He3 = 0
        .F21Ne0 = 1
        .F21Ne1 = 0
        .F21Ne2 = 0
        .F21Ne3 = 0
    Select Case equation
    Case Is = "Granger"
        .F26Al0 = 0.9655
        .F26Al1 = 0.0233
        .F26Al2 = 0.005
        .F26Al3 = 0.0062
        .F10Be0 = 0.9724
        .F10Be1 = 0.0186
        .F10Be2 = 0.004
        .F10Be3 = 0.005
        .F36Cl0 = 0.903
        .F36Cl1 = 0.0447
        .F36Cl2 = 0.0523
        .F36Cl3 = 0
        .F14C0 = 0.83
        .F14C1 = 0.0691
        .F14C2 = 0.0809
        .F14C3 = 0.02
    Case Is = "Spallation only"
        .F26Al0 = 1
        .F26Al1 = 0
        .F26Al2 = 0
        .F26Al3 = 0
        .F10Be0 = 1
        .F10Be1 = 0
        .F10Be2 = 0
        .F10Be3 = 0
        .F36Cl0 = 1
        .F36Cl1 = 0
        .F36Cl2 = 0
        .F36Cl3 = 0
        .F36Cl3 = 0
        .F14C0 = 1
        .F14C1 = 0
        .F14C2 = 0
        .F14C3 = 0
    Case Is = "Schaller"
        .F26Al0 = 0.9575
        .F26Al1 = 0.0315
        .F26Al2 = -0.009
        .F26Al3 = 0.02
        .F10Be0 = 0.964
        .F10Be1 = 0.0266
        .F10Be2 = -0.0074
        .F10Be3 = 0.0168
        .F36Cl0 = 0.903
        .F36Cl1 = 0.0793
        .F36Cl2 = 0.0177
        .F36Cl3 = 0
        .F14C0 = 0.83
        .F14C1 = 0.1363
        .F14C2 = 0.0137
        .F14C3 = 0.02
    Case Is = "Braucher"
        .L0 = 160
        .L1 = 1500
        .L2 = 1500
        .L3 = 4320
        .F26Al0 = 0.9699
        .F26Al1 = 0.0275
        .F26Al2 = 0
        .F26Al3 = 0.0026
        .F10Be0 = 0.9887
        .F10Be1 = 0.0027
        .F10Be2 = 0
        .F10Be3 = 0.0086
        .F36Cl0 = 0.9456
        .F36Cl1 = 0.0324
        .F36Cl2 = 0
        .F36Cl3 = 0.022
        .F14C0 = 0.83
        .F14C1 = 0.15
        .F14C2 = 0
        .F14C3 = 0.02
    End Select
    .equation = equation
    End With
End Sub
Public Sub reset()
    With Me
        If .PC Then
            .detail = True
            .PlotEllipse = True
        Else
            .detail = False
            .PlotEllipse = False
        End If
        .zeroerosion = False
        .AlBeLogOrLin = "log"
        .NeBeLogOrLin = "log"
        .Replace = False
        .NewtonOption = True
        .MetropIter = 1000
        .ConfiLevel = 5
        .sigma = 2
        .R_d = 287.05
        .G_o = 9.80665
        .Zero = 0.0000000001
        .yMax = 0
        .yMin = 0
        .xMin = 0
        .xMax = 0
    End With
    Call reset3
    Call reset2
    Call reset1
End Sub
Public Sub reset1()
    With Me
        .Scaling = "Lal"
        .P21Ne10Be = 4.23 'from Kober et al. (2011, EPSL, v.302, p.163-171)
        .n10BeCals = loadCalibrations("10Be", .name)
        .n26AlCals = loadCalibrations("26Al", .name)
        .n3HeCals = loadCalibrations("3He", .name)
        .n21NeCals = loadCalibrations("21Ne", .name)
        .n14CCals = loadCalibrations("14C", .name)
        .n36ClCals = loadCalibrations("36Cl", .name)
        .tieNe2Be = True
        Call setSLHLp("10Be")
        Call setSLHLp("26Al")
        Call setSLHLp("36Cl")
        Call setSLHLp("14C")
        Call setSLHLp("3He")
        Call setSLHLp("21Ne")
    End With
End Sub
Public Sub reset2()
    With Me
        .rho = 2.65
        .L3He = 0
        .L10Be = 0.0000005
        .L14C = 0.0001213
        .L21Ne = 0
        .L26Al = 0.00000098
        .L36Cl = 0.0000023
        .equation = "Braucher"
        Call setSLHLf(.equation)
    End With
End Sub
Public Sub reset3()
    With Me
        .T_o = 288.15
        .B_o = 0.0065
        .P_o = 1013.25
        .MM0 = 1
        .exponent = 2.3
    End With
End Sub
Property Get RecRange(ByVal nucl As String, ByVal recnum As Integer) As String
    Select Case nucl
    Case Is = "10Be"
        RecRange = "B" & recnum + 1 & ":H" & recnum + 1
    Case Is = "26Al"
        RecRange = "I" & recnum + 1 & ":O" & recnum + 1
    Case Is = "21Ne"
        RecRange = "P" & recnum + 1 & ":V" & recnum + 1
    Case Is = "3He"
        RecRange = "W" & recnum + 1 & ":AC" & recnum + 1
    Case Is = "36Cl"
        RecRange = "AD" & recnum + 1 & ":AJ" & recnum + 1
    Case Is = "14C"
        RecRange = "AK" & recnum + 1 & ":AQ" & recnum + 1
    End Select
End Property
Property Get CalRange(ByVal nucl As String) As String
    Select Case nucl
    Case Is = "10Be"
        CalRange = "B2:H" & Me.n10BeCals + 1
    Case Is = "26Al"
        CalRange = "I2:O" & Me.n26AlCals + 1
    Case Is = "21Ne"
        CalRange = "P2:V" & Me.n21NeCals + 1
    Case Is = "3He"
        CalRange = "W2:AC" & Me.n3HeCals + 1
    Case Is = "36Cl"
        CalRange = "AD2:AJ" & Me.n36ClCals + 1
    Case Is = "14C"
        CalRange = "AK2:AQ" & Me.n14CCals + 1
    End Select
End Property
Property Get latRange(ByVal nucl As String) As String
    Select Case nucl
    Case Is = "10Be"
        latRange = "D2:D" & Me.n10BeCals + 1
    Case Is = "26Al"
        latRange = "K2:K" & Me.n26AlCals + 1
    Case Is = "21Ne"
        latRange = "R2:R" & Me.n21NeCals + 1
    Case Is = "3He"
        latRange = "Y2:Y" & Me.n3HeCals + 1
    Case Is = "36Cl"
        latRange = "AF2:AF" & Me.n36ClCals + 1
    Case Is = "14C"
        latRange = "AM2:AM" & Me.n14CCals + 1
    End Select
End Property
Property Get elevRange(ByVal nucl As String) As String
    Select Case nucl
    Case Is = "10Be"
        elevRange = "E2:E" & Me.n10BeCals + 1
    Case Is = "26Al"
        elevRange = "L2:L" & Me.n26AlCals + 1
    Case Is = "21Ne"
        elevRange = "S2:S" & Me.n21NeCals + 1
    Case Is = "3He"
        elevRange = "Z2:Z" & Me.n3HeCals + 1
    Case Is = "36Cl"
        elevRange = "AG2:AG" & Me.n36ClCals + 1
    Case Is = "14C"
        elevRange = "AN2:AN" & Me.n14CCals + 1
    End Select
End Property
Property Get LatElevRange(ByVal nucl As String) As String
    Select Case nucl
    Case Is = "10Be"
        LatElevRange = "D2:E" & Me.n10BeCals + 1
    Case Is = "26Al"
        LatElevRange = "K2:L" & Me.n26AlCals + 1
    Case Is = "21Ne"
        LatElevRange = "R2:S" & Me.n21NeCals + 1
    Case Is = "3He"
        LatElevRange = "Y2:Z" & Me.n3HeCals + 1
    Case Is = "36Cl"
        LatElevRange = "AF2:AG" & Me.n36ClCals + 1
    Case Is = "14C"
        LatElevRange = "AM2:AN" & Me.n14CCals + 1
    End Select
End Property
Property Let name(ByVal newvalue As String)
    m_name = newvalue
End Property
Property Get name() As String
    name = m_name
End Property
Property Let version(ByVal newvalue As String)
    m_version = newvalue
    Worksheets(m_name).Range("version").Value = m_version
End Property
Property Get version() As String
    version = m_version
End Property
Property Let rho(ByVal newvalue As Double)
    m_rho = newvalue
    Worksheets(m_name).Range("rho").Value = m_rho
End Property
Property Get rho() As Double
    rho = m_rho
End Property
Property Get PC() As Boolean
    PC = m_PC
End Property
Property Let detail(ByVal newvalue As Double)
    m_detail = newvalue
    Worksheets(m_name).Range("detail").Value = m_detail
End Property
Property Get detail() As Double
    detail = m_detail
End Property
Property Let zeroerosion(ByVal newvalue As Double)
    m_zeroerosion = newvalue
    Worksheets(m_name).Range("zeroerosion").Value = m_zeroerosion
End Property
Property Get zeroerosion() As Double
    zeroerosion = m_zeroerosion
End Property
Property Let L3He(ByVal newvalue As Double)
    m_L3He = newvalue
    Worksheets(m_name).Range("L_3He").Value = m_L3He
End Property
Property Get L3He() As Double
    L3He = m_L3He
End Property
Property Let L10Be(ByVal newvalue As Double)
    m_L10Be = newvalue
    Worksheets(m_name).Range("L_10Be").Value = m_L10Be
End Property
Property Get L10Be() As Double
    L10Be = m_L10Be
End Property
Property Let L14C(ByVal newvalue As Double)
    m_L14C = newvalue
    Worksheets(m_name).Range("L_14C").Value = m_L14C
End Property
Property Get L14C() As Double
    L14C = m_L14C
End Property
Property Let L21Ne(ByVal newvalue As Double)
    m_L21Ne = newvalue
    Worksheets(m_name).Range("L_21Ne").Value = m_L21Ne
End Property
Property Get L21Ne() As Double
    L21Ne = m_L21Ne
End Property
Property Let L26Al(ByVal newvalue As Double)
    m_L26Al = newvalue
    Worksheets(m_name).Range("L_26Al").Value = m_L26Al
End Property
Property Get L26Al() As Double
    L26Al = m_L26Al
End Property
Property Let L36Cl(ByVal newvalue As Double)
    m_L36Cl = newvalue
    Worksheets(m_name).Range("L_36Cl").Value = m_L36Cl
End Property
Property Get L36Cl() As Double
    L36Cl = m_L36Cl
End Property
Property Let n10BeCals(ByVal newvalue As Integer)
    m_n10BeCals = newvalue
    Worksheets(m_name).Range("n_10BeCals").Value = m_n10BeCals
End Property
Property Get n10BeCals() As Integer
    n10BeCals = m_n10BeCals
End Property
Property Let n26AlCals(ByVal newvalue As Integer)
    m_n26AlCals = newvalue
    Worksheets(m_name).Range("n_26AlCals").Value = m_n26AlCals
End Property
Property Get n26AlCals() As Integer
    n26AlCals = m_n26AlCals
End Property
Property Let n21NeCals(ByVal newvalue As Integer)
    m_n21NeCals = newvalue
    Worksheets(m_name).Range("n_21NeCals").Value = m_n21NeCals
End Property
Property Get n21NeCals() As Integer
    n21NeCals = m_n21NeCals
End Property
Property Let n3HeCals(ByVal newvalue As Integer)
    m_n3HeCals = newvalue
    Worksheets(m_name).Range("n_3HeCals").Value = m_n3HeCals
End Property
Property Get n3HeCals() As Integer
    n3HeCals = m_n3HeCals
End Property
Property Let n36ClCals(ByVal newvalue As Integer)
    m_n36ClCals = newvalue
    Worksheets(m_name).Range("n_36ClCals").Value = m_n36ClCals
End Property
Property Get n36ClCals() As Integer
    n36ClCals = m_n36ClCals
End Property
Property Let n14CCals(ByVal newvalue As Integer)
    m_n14CCals = newvalue
    Worksheets(m_name).Range("n_14CCals").Value = m_n14CCals
End Property
Property Get n14CCals() As Integer
    n14CCals = m_n14CCals
End Property
Property Let AlBeLogOrLin(ByVal newvalue As String)
    m_AlBeLogOrLin = newvalue
    Worksheets(m_name).Range("AlBeLogOrLin").Value = m_AlBeLogOrLin
End Property
Property Get AlBeLogOrLin() As String
    AlBeLogOrLin = m_AlBeLogOrLin
End Property
Property Let NeBeLogOrLin(ByVal newvalue As String)
    m_NeBeLogOrLin = newvalue
    Worksheets(m_name).Range("NeBeLogOrLin").Value = m_NeBeLogOrLin
End Property
Property Get NeBeLogOrLin() As String
    NeBeLogOrLin = m_NeBeLogOrLin
End Property
Property Let Replace(ByVal newvalue As Boolean)
    m_Replace = newvalue
    Worksheets(m_name).Range("Replace").Value = m_Replace
End Property
Property Get Replace() As Boolean
    Replace = m_Replace
End Property
Property Let PlotEllipse(ByVal newvalue As Boolean)
    m_PlotEllipse = newvalue
    Worksheets(m_name).Range("PlotEllipse").Value = m_PlotEllipse
End Property
Property Get PlotEllipse() As Boolean
    PlotEllipse = m_PlotEllipse
End Property
Property Let NewtonOption(ByVal newvalue As Boolean)
    m_NewtonOption = newvalue
    Worksheets(m_name).Range("NewtonOption").Value = m_NewtonOption
End Property
Property Get NewtonOption() As Boolean
    NewtonOption = m_NewtonOption
End Property
Property Let MetropIter(ByVal newvalue As Integer)
    m_MetropIter = newvalue
    Worksheets(m_name).Range("MetropIter").Value = m_MetropIter
End Property
Property Get MetropIter() As Integer
    MetropIter = m_MetropIter
End Property
Property Let ConfiLevel(ByVal newvalue As Double)
    m_ConfiLevel = newvalue
    Worksheets(m_name).Range("ConfiLevel").Value = m_ConfiLevel
End Property
Property Get ConfiLevel() As Double
    ConfiLevel = m_ConfiLevel
End Property
Property Let T_o(ByVal newvalue As Double)
    m_T_o = newvalue
    Worksheets(m_name).Range("T_o").Value = m_T_o
End Property
Property Get T_o() As Double
    T_o = m_T_o
End Property
Property Let B_o(ByVal newvalue As Double)
    m_B_o = newvalue
    Worksheets(m_name).Range("B_o").Value = m_B_o
End Property
Property Get B_o() As Double
    B_o = m_B_o
End Property
Property Let R_d(ByVal newvalue As Double)
    m_R_d = newvalue
    Worksheets(m_name).Range("R_d").Value = m_R_d
End Property
Property Get R_d() As Double
    R_d = m_R_d
End Property
Property Let G_o(ByVal newvalue As Double)
    m_G_o = newvalue
    Worksheets(m_name).Range("G_o").Value = m_G_o
End Property
Property Get G_o() As Double
    G_o = m_G_o
End Property
Property Let P_o(ByVal newvalue As Double)
    m_P_o = newvalue
    Worksheets(m_name).Range("P_o").Value = m_P_o
End Property
Property Get P_o() As Double
    P_o = m_P_o
End Property
Property Let MM0(ByVal newvalue As Double)
    m_MM0 = newvalue
    Worksheets(m_name).Range("MM_0").Value = m_MM0
End Property
Property Get MM0() As Double
    MM0 = m_MM0
End Property
Property Let Zero(ByVal newvalue As Double)
    m_Zero = newvalue
    Worksheets(m_name).Range("Zero").Value = m_Zero
End Property
Property Get Zero() As Double
    Zero = m_Zero
End Property
Property Let L0(ByVal newvalue As Double)
    m_L0 = newvalue
    Worksheets(m_name).Range("L_0").Value = m_L0
End Property
Property Get L0() As Double
    L0 = m_L0
End Property
Property Let L1(ByVal newvalue As Double)
    m_L1 = newvalue
    Worksheets(m_name).Range("L_1").Value = m_L1
End Property
Property Get L1() As Double
    L1 = m_L1
End Property
Property Let L2(ByVal newvalue As Double)
    m_L2 = newvalue
    Worksheets(m_name).Range("L_2").Value = m_L2
End Property
Property Get L2() As Double
    L2 = m_L2
End Property
Property Let L3(ByVal newvalue As Double)
    m_L3 = newvalue
    Worksheets(m_name).Range("L_3").Value = m_L3
End Property
Property Get L3() As Double
    L3 = m_L3
End Property
Property Let P3He(ByVal newvalue As Double)
    m_P3He = newvalue
    Worksheets(m_name).Range("P_3He").Value = m_P3He
End Property
Property Get P3He() As Double
    P3He = m_P3He
End Property
Property Let F3He0(ByVal newvalue As Double)
    m_F3He0 = newvalue
    Worksheets(m_name).Range("F_3He0").Value = m_F3He0
End Property
Property Get F3He0() As Double
    F3He0 = m_F3He0
End Property
Property Let F3He1(ByVal newvalue As Double)
    m_F3He1 = newvalue
    Worksheets(m_name).Range("F_3He1").Value = m_F3He1
End Property
Property Get F3He1() As Double
    F3He1 = m_F3He1
End Property
Property Let F3He2(ByVal newvalue As Double)
    m_F3He2 = newvalue
    Worksheets(m_name).Range("F_3He2").Value = m_F3He2
End Property
Property Get F3He2() As Double
    F3He2 = m_F3He2
End Property
Property Let F3He3(ByVal newvalue As Double)
    m_F3He3 = newvalue
    Worksheets(m_name).Range("F_3He3").Value = m_F3He3
End Property
Property Get F3He3() As Double
    F3He3 = m_F3He3
End Property
Property Let P21Ne(ByVal newvalue As Double)
    m_P21Ne = newvalue
    Worksheets(m_name).Range("P_21Ne").Value = m_P21Ne
End Property
Property Get P21Ne() As Double
    P21Ne = m_P21Ne
End Property
Property Let F21Ne0(ByVal newvalue As Double)
    m_F21Ne0 = newvalue
    Worksheets(m_name).Range("F_21Ne0").Value = m_F21Ne0
End Property
Property Get F21Ne0() As Double
    F21Ne0 = m_F21Ne0
End Property
Property Let F21Ne1(ByVal newvalue As Double)
    m_F21Ne1 = newvalue
    Worksheets(m_name).Range("F_21Ne1").Value = m_F21Ne1
End Property
Property Get F21Ne1() As Double
    F21Ne1 = m_F21Ne1
End Property
Property Let F21Ne2(ByVal newvalue As Double)
    m_F21Ne2 = newvalue
    Worksheets(m_name).Range("F_21Ne2").Value = m_F21Ne2
End Property
Property Get F21Ne2() As Double
    F21Ne2 = m_F21Ne2
End Property
Property Let F21Ne3(ByVal newvalue As Double)
    m_F21Ne3 = newvalue
    Worksheets(m_name).Range("F_21Ne3").Value = m_F21Ne3
End Property
Property Get F21Ne3() As Double
    F21Ne3 = m_F21Ne3
End Property
Property Let P14C(ByVal newvalue As Double)
    m_P14C = newvalue
    Worksheets(m_name).Range("P_14C").Value = m_P14C
End Property
Property Get P14C() As Double
    P14C = m_P14C
End Property
Property Let F14C0(ByVal newvalue As Double)
    m_F14C0 = newvalue
    Worksheets(m_name).Range("F_14C0").Value = m_F14C0
End Property
Property Get F14C0() As Double
    F14C0 = m_F14C0
End Property
Property Let F14C1(ByVal newvalue As Double)
    m_F14C1 = newvalue
    Worksheets(m_name).Range("F_14C1").Value = m_F14C1
End Property
Property Get F14C1() As Double
    F14C1 = m_F14C1
End Property
Property Let F14C2(ByVal newvalue As Double)
    m_F14C2 = newvalue
    Worksheets(m_name).Range("F_14C2").Value = m_F14C2
End Property
Property Get F14C2() As Double
    F14C2 = m_F14C2
End Property
Property Let F14C3(ByVal newvalue As Double)
    m_F14C3 = newvalue
    Worksheets(m_name).Range("F_14C3").Value = m_F14C3
End Property
Property Get F14C3() As Double
    F14C3 = m_F14C3
End Property
Property Let P26Al(ByVal newvalue As Double)
    m_P26Al = newvalue
    Worksheets(m_name).Range("P_26Al").Value = m_P26Al
End Property
Property Get P26Al() As Double
    P26Al = m_P26Al
End Property
Property Let F26Al0(ByVal newvalue As Double)
    m_F26Al0 = newvalue
    Worksheets(m_name).Range("F_26Al0").Value = m_F26Al0
End Property
Property Get F26Al0() As Double
    F26Al0 = m_F26Al0
End Property
Property Let F26Al1(ByVal newvalue As Double)
    m_F26Al1 = newvalue
    Worksheets(m_name).Range("F_26Al1").Value = m_F26Al1
End Property
Property Get F26Al1() As Double
    F26Al1 = m_F26Al1
End Property
Property Let F26Al2(ByVal newvalue As Double)
    m_F26Al2 = newvalue
    Worksheets(m_name).Range("F_26Al2").Value = m_F26Al2
End Property
Property Get F26Al2() As Double
    F26Al2 = m_F26Al2
End Property
Property Let F26Al3(ByVal newvalue As Double)
    m_F26Al3 = newvalue
    Worksheets(m_name).Range("F_26Al3").Value = m_F26Al3
End Property
Property Get F26Al3() As Double
    F26Al3 = m_F26Al3
End Property
Property Let P36Cl(ByVal newvalue As Double)
    m_P36Cl = newvalue
    Worksheets(m_name).Range("P_36Cl").Value = m_P36Cl
End Property
Property Get P36Cl() As Double
    P36Cl = m_P36Cl
End Property
Property Let F36Cl0(ByVal newvalue As Double)
    m_F36Cl0 = newvalue
    Worksheets(m_name).Range("F_36Cl0").Value = m_F36Cl0
End Property
Property Get F36Cl0() As Double
    F36Cl0 = m_F36Cl0
End Property
Property Let F36Cl1(ByVal newvalue As Double)
    m_F36Cl1 = newvalue
    Worksheets(m_name).Range("F_36Cl1").Value = m_F36Cl1
End Property
Property Get F36Cl1() As Double
    F36Cl1 = m_F36Cl1
End Property
Property Let F36Cl2(ByVal newvalue As Double)
    m_F36Cl2 = newvalue
    Worksheets(m_name).Range("F_36Cl2").Value = m_F36Cl2
End Property
Property Get F36Cl2() As Double
    F36Cl2 = m_F36Cl2
End Property
Property Let F36Cl3(ByVal newvalue As Double)
    m_F36Cl3 = newvalue
    Worksheets(m_name).Range("F_36Cl3").Value = m_F36Cl3
End Property
Property Get F36Cl3() As Double
    F36Cl3 = m_F36Cl3
End Property
Property Let P10Be(ByVal newvalue As Double)
    m_P10Be = newvalue
    Worksheets(m_name).Range("P_10Be").Value = m_P10Be
End Property
Property Get P10Be() As Double
    P10Be = m_P10Be
End Property
Property Let F10Be0(ByVal newvalue As Double)
    m_F10Be0 = newvalue
    Worksheets(m_name).Range("F_10Be0").Value = m_F10Be0
End Property
Property Get F10Be0() As Double
    F10Be0 = m_F10Be0
End Property
Property Let F10Be1(ByVal newvalue As Double)
    m_F10Be1 = newvalue
    Worksheets(m_name).Range("F_10Be1").Value = m_F10Be1
End Property
Property Get F10Be1() As Double
    F10Be1 = m_F10Be1
End Property
Property Let F10Be2(ByVal newvalue As Double)
    m_F10Be2 = newvalue
    Worksheets(m_name).Range("F_10Be2").Value = m_F10Be2
End Property
Property Get F10Be2() As Double
    F10Be2 = m_F10Be2
End Property
Property Let F10Be3(ByVal newvalue As Double)
    m_F10Be3 = newvalue
    Worksheets(m_name).Range("F_10Be3").Value = m_F10Be3
End Property
Property Get F10Be3() As Double
    F10Be3 = m_F10Be3
End Property
Property Let P21Ne10Be(ByVal newvalue As Double)
    m_P21Ne10Be = newvalue
    Worksheets(m_name).Range("P_21Ne10Be").Value = m_P21Ne10Be
End Property
Property Get P21Ne10Be() As Double
    P21Ne10Be = m_P21Ne10Be
End Property
Property Let Scaling(ByVal newvalue As String)
    m_scaling = newvalue
    Worksheets(m_name).Range("scaling").Value = m_scaling
End Property
Property Get Scaling() As String
    Scaling = m_scaling
End Property
Property Let equation(ByVal newvalue As String)
    m_equation = newvalue
    Worksheets(m_name).Range("equation").Value = m_equation
End Property
Property Get equation() As String
    equation = m_equation
End Property
Property Let yMax(newvalue As Double)
    m_yMax = newvalue
    Worksheets(m_name).Range("yMax").Value = m_yMax
End Property
Property Get yMax() As Double
    yMax = m_yMax
End Property
Property Let yMin(newvalue As Double)
    m_yMin = newvalue
    Worksheets(m_name).Range("yMin").Value = m_yMin
End Property
Property Get yMin() As Double
    yMin = m_yMin
End Property
Property Let xMin(newvalue As Double)
    m_xMin = newvalue
    Worksheets(m_name).Range("xMin").Value = m_xMin
End Property
Property Get xMin() As Double
    xMin = m_xMin
End Property
Property Let xMax(newvalue As Double)
    m_xMax = newvalue
    Worksheets(m_name).Range("xMax").Value = m_xMax
End Property
Property Get xMax() As Double
    xMax = m_xMax
End Property
Property Let exponent(newvalue As Double)
    m_exponent = newvalue
    Worksheets(m_name).Range("exponent").Value = m_exponent
End Property
Property Get exponent() As Double
    exponent = m_exponent
End Property
Property Let sigma(newvalue As Double)
    m_sigma = newvalue
    Worksheets(m_name).Range("sigma").Value = m_sigma
End Property
Property Get sigma() As Double
    sigma = m_sigma
End Property
Property Let tieNe2Be(newvalue As Boolean)
    m_tieNe2Be = newvalue
    Worksheets(m_name).Range("tieNe2Be").Value = m_tieNe2Be
End Property
Property Get tieNe2Be() As Boolean
    tieNe2Be = m_tieNe2Be
End Property
