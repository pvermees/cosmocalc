Attribute VB_Name = "Module_AgeCalc"
Public Function getAge(ByVal N As Double, nucl As MyNuclide, ByVal E As Double) As Double
    Dim t As Double
    Dim dNdt As Double
    P = nucl.P * (nucl.S0 * nucl.F0 + nucl.S1 * nucl.F1 + nucl.S2 * nucl.F2 + nucl.F3 * nucl.F3)
    t = N / P ' initialize
    If (E = 0) And (nucl.L <> 0) Then
        t = -(1 / nucl.L) * Log(1 - N * nucl.L / P)
    End If
    If (E <> 0) Then
        For i = 1 To 100
            Nest = getN(E, t, 0, nucl)
            If Abs(Nest - N) < 1 Then
                Exit For
            End If
            dNdt = nucl.P * nucl.S0 * nucl.F0 * Exp(-t * (nucl.L + E * glob.rho / glob.L0)) + _
                   nucl.P * nucl.S1 * nucl.F1 * Exp(-t * (nucl.L + E * glob.rho / glob.L1)) + _
                   nucl.P * nucl.S2 * nucl.F2 * Exp(-t * (nucl.L + E * glob.rho / glob.L2)) + _
                   nucl.P * nucl.S3 * nucl.F3 * Exp(-t * (nucl.L + E * glob.rho / glob.L3))
            dt = (Nest - N) / dNdt
            t = t - dt
        Next i
    End If
    getAge = t
End Function
Public Function getAgeErr(ByVal N As Double, ByVal Nerr As Double, nucl As MyNuclide, ByVal t As Double, ByVal E As Double) As Double
    dN = dNdt(nucl, E, t, 0)
    getAgeErr = Nerr / dN
End Function
Public Function getErosion(ByVal N As Double, nucl As MyNuclide) As Double
    Dim dNdE As Double
    P = nucl.P * (nucl.S0 * nucl.F0 + nucl.S1 * nucl.F1 + nucl.S2 * nucl.F2 + nucl.F3 * nucl.F3)
    E = (glob.L0 / glob.rho) * ((P / N) - nucl.L)
    ' find root with Newton's method
    For i = 1 To 100
        Nest = getN(E, "inf", 0, nucl)
        If Abs(Nest - N) < 1 Then
            Exit For
        End If
        dNdE = nucl.P * nucl.S0 * nucl.F0 * (glob.rho / glob.L0) / (nucl.L + E * glob.rho / glob.L0) ^ 2 + _
               nucl.P * nucl.S1 * nucl.F1 * (glob.rho / glob.L1) / (nucl.L + E * glob.rho / glob.L1) ^ 2 + _
               nucl.P * nucl.S2 * nucl.F2 * (glob.rho / glob.L2) / (nucl.L + E * glob.rho / glob.L2) ^ 2 + _
               nucl.P * nucl.S3 * nucl.F3 * (glob.rho / glob.L3) / (nucl.L + E * glob.rho / glob.L3) ^ 2
        dE = (Nest - N) / dNdE
        E = E + dE
    Next i
    getErosion = E
End Function
Public Function getErosionErr(ByVal N As Double, ByVal Nerr As Double, nucl As MyNuclide) As Double
    Dim dNdE As Double
    Dim dNdt As Double
    E = getErosion(N, nucl)
    dNdE = nucl.P * nucl.S0 * nucl.F0 / ((glob.L0 / glob.rho) * (nucl.L + E / (glob.L0 / glob.rho)) ^ 2) + _
           nucl.P * nucl.S1 * nucl.F1 / ((glob.L1 / glob.rho) * (nucl.L + E / (glob.L1 / glob.rho)) ^ 2) + _
           nucl.P * nucl.S2 * nucl.F2 / ((glob.L2 / glob.rho) * (nucl.L + E / (glob.L2 / glob.rho)) ^ 2) + _
           nucl.P * nucl.S3 * nucl.F3 / ((glob.L3 / glob.rho) * (nucl.L + E / (glob.L3 / glob.rho)) ^ 2)
    getErosionErr = Nerr / dNdE
End Function
Public Function getBurial(ByVal N As Double, nucl As MyNuclide, ByVal E As Double, ByVal t As Variant) As Double
    ' get the burial age given the erosion rate and exposure age
    If t = "inf" Then
        getBurial = (1 / nucl.L) * Log((1 / N) * ( _
                    (nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0)) + _
                    (nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1)) + _
                    (nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2)) + _
                    (nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3))))
    Else
        getBurial = (1 / nucl.L) * Log((1 / N) * ( _
                    (nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0)) * (1 - Exp(-(nucl.L + E * glob.rho / glob.L0) * t)) + _
                    (nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1)) * (1 - Exp(-(nucl.L + E / (glob.L1 / glob.rho)) * t)) + _
                    (nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2)) * (1 - Exp(-(nucl.L + E / (glob.L2 / glob.rho)) * t)) + _
                    (nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3)) * (1 - Exp(-(nucl.L + E / (glob.L3 / glob.rho)) * t))))
    End If
End Function
Public Function getBurialErr(ByVal N As Double, ByVal Nerr As Double, nucl As MyNuclide, ByVal E As Double, ByVal Eerr As Double) As Double
    ' get the burial age given the steady state erosion rate
    Dim dTaudN As Double
    Dim dTaudE As Double
    dTaudN = -(1 / (N * nucl.L)) * (1 / ( _
            nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0) + _
            nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1) + _
            nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2) + _
            nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3) _
            ))
    dTaudE = -(1 / nucl.L) * (1 / ( _
            nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0) + _
            nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1) + _
            nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2) + _
            nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3) _
            )) * ( _
            (nucl.P * nucl.S0 * nucl.F0 * glob.rho / glob.L0) / (nucl.L + E * glob.rho / glob.L0) ^ 2 + _
            (nucl.P * nucl.S1 * nucl.F1 * glob.rho / glob.L1) / (nucl.L + E * glob.rho / glob.L1) ^ 2 + _
            (nucl.P * nucl.S2 * nucl.F2 * glob.rho / glob.L2) / (nucl.L + E * glob.rho / glob.L2) ^ 2 + _
            (nucl.P * nucl.S3 * nucl.F3 * glob.rho / glob.L3) / (nucl.L + E * glob.rho / glob.L3) ^ 2 _
            )
    getBurialErr = ((dTaudN ^ 2) * (Nerr ^ 2) + (dTaudE ^ 2) * (Eerr ^ 2)) ^ 0.5
End Function
Public Function getBurialErr2(ByVal N As Double, ByVal Nerr As Double, nucl As MyNuclide, ByVal t As Double, ByVal terr As Double) As Double
    ' get the burial age given the steady state erosion rate
    Dim dTaudN As Double
    Dim dTaudt As Double
    dTaudN = -1 / (N * nucl.L)
    dTaudt = Exp(-nucl.L * t) / (1 - Exp(-nucl.L * t))
    getBurialErr2 = ((dTaudN ^ 2) * (Nerr ^ 2) + (dTaudt ^ 2) * (terr ^ 2)) ^ 0.5
End Function
Public Function getN(ByVal E As Double, ByVal t As Variant, ByVal burialT As Double, nucl As MyNuclide) As Double
    If (t <> "inf") And (E <> 0 Or nucl.L <> 0) Then
        getN = ((nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0)) * (1 - Exp(-(nucl.L + E * glob.rho / glob.L0) * t)) + _
                (nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1)) * (1 - Exp(-(nucl.L + E * glob.rho / glob.L1) * t)) + _
                (nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2)) * (1 - Exp(-(nucl.L + E * glob.rho / glob.L2) * t)) + _
                (nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3)) * (1 - Exp(-(nucl.L + E * glob.rho / glob.L3) * t))) * Exp(-burialT * nucl.L)
    ElseIf t = "inf" Then
        getN = (nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0) + _
                nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1) + _
                nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2) + _
                nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3)) * Exp(-burialT * nucl.L)
    ElseIf (E = 0 And nucl.L = 0) Then
        getN = t * nucl.P * (nucl.S0 * nucl.F0 + nucl.S1 * nucl.F1 + nucl.S2 * nucl.F2 + nucl.S3 * nucl.F3)
    End If
End Function
Public Sub getBurialErosion(ByVal N1 As Double, ByVal N1err As Double, ByVal N2 As Double, ByVal N2err As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByRef E As Double, ByRef Eerr As Double, ByRef Tau As Double, ByRef TauErr As Double)
    Dim a As Double, B As Double, c As Double, d As Double
    On Error GoTo errorhandler:
    ' initialize E
    If nuclide1.L < nuclide2.L Then
        E = getErosion(N1, nuclide1)
    Else
        E = getErosion(N2, nuclide2)
    End If
    ' Newton's method
    For j = 1 To 100
        F = fBe(N1, N2, E, nuclide1, nuclide2)
        dFdE = DfBeDe(N1, N2, E, nuclide1, nuclide2)
        dE = F / dFdE
        E = Abs(E - dE)
        If Abs(dE / E) < glob.Zero Then
            Exit For
        End If
    Next j
    If (nuclide1.L > nuclide2.L) Then
        Tau = getBurial(N1, nuclide1, E, "inf")
    Else
        Tau = getBurial(N2, nuclide2, E, "inf")
    End If
    Call getJacobian("burial_erosion", nuclide1, nuclide2, a, B, c, d, E, Tau)
    Eerr = (Abs(B * N1err) + Abs(d * N2err)) / Abs(a * d - B * c)
    TauErr = (Abs(a * N1err) + Abs(c * N2err)) / Abs(a * d - B * c)
Label:
errorhandler:
    If (Err.Number <> 0) Then
        E = 0
        Eerr = 0
        Tau = 0
        TauErr = 0
        Resume Label
    End If
End Sub
Public Sub getBurialExposure(ByVal N1 As Double, ByVal N1err As Double, ByVal N2 As Double, ByVal N2err As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByRef t As Double, ByRef terr As Double, ByRef Tau As Double, ByRef TauErr As Double)
    Dim a As Double, B As Double, c As Double, d As Double
    On Error GoTo errorhandler:
    ' initialize t
    If nuclide1.L < nuclide2.L Then
        t = getAge(N1, nuclide1, 0)
    Else
        t = getAge(N2, nuclide2, 0)
    End If
    If nuclide1.L = 0 Then
        Tau = getBurial(N2, nuclide2, 0, t)
        terr = getAgeErr(N1, N1err, nuclide1, t, 0)
        TauErr = getBurialErr2(N2, N2err, nuclide2, t, terr)
    ElseIf nuclide2.L = 0 Then
        Tau = getBurial(N1, nuclide1, 0, t)
        terr = getAgeErr(N2, N2err, nuclide2, t, 0)
        TauErr = getBurialErr2(N1, N1err, nuclide1, t, terr)
    Else
        ' Newton's method
        For j = 1 To 100
            F = fBt(N1, N2, t, nuclide1, nuclide2)
            dFdT = DfBtDt(N1, N2, t, nuclide1, nuclide2)
            dt = F / dFdT
            t = Abs(t - dt)
            If Abs(dt / t) < glob.Zero Then
                Exit For
            End If
        Next j
        If (nuclide1.L > nuclide2.L) Then
            Tau = getBurial(N1, nuclide1, 0, t)
        Else
            Tau = getBurial(N2, nuclide2, 0, t)
        End If
        Call getJacobian("burial_exposure", nuclide1, nuclide2, a, B, c, d, t, Tau)
        terr = (Abs(B * N1err) + Abs(d * N2err)) / Abs(a * d - B * c)
        TauErr = (Abs(a * N1err) + Abs(c * N2err)) / Abs(a * d - B * c)
    End If
Label:
errorhandler:
    If (Err.Number <> 0) Then
        t = 0
        terr = 0
        Tau = 0
        TauErr = 0
        Resume Label
    End If
End Sub
Public Sub getAgeErosion(ByVal N1 As Double, ByVal N1err As Double, ByVal N2 As Double, ByVal N2err As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByRef E As Double, ByRef Eerr As Double, ByRef t As Double, ByRef terr As Double)
    Dim a As Double, B As Double, c As Double, d As Double
    On Error GoTo errorhandler:
    ' initialize t and E
    minE = glob.Zero 'cm/yr
    maxE = 0.1 'cm/yr
    mint = 1000 'yr
    maxt = 20000000 'yr
    If nuclide1.L > nuclide2.L Then
        t = 1 / nuclide1.L
        E = getErosion(N2, nuclide2)
    Else
        t = 1 / nuclide2.L
        E = getErosion(N1, nuclide1)
    End If
    ' two dimensional Newton
    For j = 1 To 100
        N1est = getN(E, t, 0, nuclide1)
        N2est = getN(E, t, 0, nuclide2)
        F1 = N1est - N1
        F2 = N2est - N2
        If (Abs(F1) < 1) And (Abs(F2) < 1) Then
            Exit For
        End If
        Call getJacobian("age_erosion", nuclide1, nuclide2, a, B, c, d, E, t)
        dE = (-d * F1 + B * F2) / (a * d - B * c)
        dt = (c * F1 - a * F2) / (a * d - B * c)
        E = Abs(E + dE)
        t = Abs(t + dt)
    Next j
    Call getJacobian("age_erosion", nuclide1, nuclide2, a, B, c, d, E, t)
    Eerr = (Abs(B * N1err) + Abs(d * N2err)) / Abs(a * d - B * c)
    terr = (Abs(a * N1err) + Abs(c * N2err)) / Abs(a * d - B * c)
errorhandler:
    If (Err.Number <> 0) Then
        'MsgBox "Algorithm did not converge", vbExclamation
        E = 0
        t = 0
        Eerr = 0
        terr = 0
    End If
End Sub
Private Function dNdTau(nucl As MyNuclide, ByVal E As Double, ByVal t As Variant, ByVal Tau As Double) As Double
    If t = "inf" Then
        dNdTau = -nucl.L * Exp(-nucl.L * Tau) * _
             ( _
                nucl.P * nucl.S0 * nucl.F0 / (nucl.L + E * glob.rho / glob.L0) + _
                nucl.P * nucl.S1 * nucl.F1 / (nucl.L + E * glob.rho / glob.L1) + _
                nucl.P * nucl.S2 * nucl.F2 / (nucl.L + E * glob.rho / glob.L2) + _
                nucl.P * nucl.S3 * nucl.F3 / (nucl.L + E * glob.rho / glob.L3) _
             )
    Else
        dNdTau = -nucl.L * Exp(-nucl.L * Tau) * _
             ( _
                nucl.P * nucl.S0 * nucl.F0 * (1 - Exp(-(nucl.L + E * glob.rho / glob.L0) * t)) / (nucl.L + E * glob.rho / glob.L0) + _
                nucl.P * nucl.S1 * nucl.F1 * (1 - Exp(-(nucl.L + E * glob.rho / glob.L1) * t)) / (nucl.L + E * glob.rho / glob.L1) + _
                nucl.P * nucl.S2 * nucl.F2 * (1 - Exp(-(nucl.L + E * glob.rho / glob.L2) * t)) / (nucl.L + E * glob.rho / glob.L2) + _
                nucl.P * nucl.S3 * nucl.F3 * (1 - Exp(-(nucl.L + E * glob.rho / glob.L3) * t)) / (nucl.L + E * glob.rho / glob.L3) _
             )
    End If
End Function
Private Function dNdE(nucl As MyNuclide, ByVal E As Double, ByVal t As Variant, ByVal Tau As Double) As Double
    If t = "inf" Then
        mu0 = glob.rho / glob.L0
        Lemu0 = nucl.L + E * mu0
        term0 = -nucl.P * nucl.S0 * nucl.F0 * mu0 / (Lemu0 ^ 2)
        mu1 = glob.rho / glob.L1
        Lemu1 = nucl.L + E * mu1
        term1 = -nucl.P * nucl.S1 * nucl.F1 * mu1 / (Lemu1 ^ 2)
        mu2 = glob.rho / glob.L2
        Lemu2 = nucl.L + E * mu2
        term2 = -nucl.P * nucl.S2 * nucl.F2 * mu2 / (Lemu2 ^ 2)
        mu3 = glob.rho / glob.L3
        Lemu3 = nucl.L + E * mu3
        term3 = -nucl.P * nucl.S3 * nucl.F3 * mu3 / (Lemu3 ^ 2)
        dNdE = Exp(-nucl.L * Tau) * (term0 + term1 + term2 + term3)
    Else
        mu0 = glob.rho / glob.L0
        Lemu0 = nucl.L + E * mu0
        term0 = (nucl.P * nucl.S0 * nucl.F0 * mu0 * t * Lemu0 * Exp(-t * Lemu0) - _
                 nucl.P * nucl.S0 * nucl.F0 * mu0 * (1 - Exp(-t * Lemu0))) / (Lemu0 ^ 2)
        mu1 = glob.rho / glob.L1
        Lemu1 = nucl.L + E * mu1
        term1 = (nucl.P * nucl.S1 * nucl.F1 * mu1 * t * Lemu1 * Exp(-t * Lemu1) - _
                 nucl.P * nucl.S1 * nucl.F1 * mu1 * (1 - Exp(-t * Lemu1))) / (Lemu1 ^ 2)
        mu2 = glob.rho / glob.L2
        Lemu2 = nucl.L + E * mu2
        term2 = (nucl.P * nucl.S2 * nucl.F2 * mu2 * t * Lemu2 * Exp(-t * Lemu2) - _
                 nucl.P * nucl.S2 * nucl.F2 * mu2 * (1 - Exp(-t * Lemu2))) / (Lemu2 ^ 2)
        mu3 = glob.rho / glob.L3
        Lemu3 = nucl.L + E * mu3
        term3 = (nucl.P * nucl.S3 * nucl.F3 * mu3 * t * Lemu3 * Exp(-t * Lemu3) - _
                 nucl.P * nucl.S3 * nucl.F3 * mu3 * (1 - Exp(-t * Lemu3))) / (Lemu3 ^ 2)
        dNdE = Exp(-nucl.L * Tau) * (term0 + term1 + term2 + term3)
    End If
End Function
Private Function dNdt(nucl As MyNuclide, ByVal E As Double, ByVal t As Variant, ByVal Tau As Double) As Double
    dNdt = Exp(-nucl.L * Tau) * _
        ( _
        nucl.P * nucl.S0 * nucl.F0 * Exp(-(nucl.L + E * glob.rho / glob.L0) * t) + _
        nucl.P * nucl.S1 * nucl.F1 * Exp(-(nucl.L + E * glob.rho / glob.L1) * t) + _
        nucl.P * nucl.S2 * nucl.F2 * Exp(-(nucl.L + E * glob.rho / glob.L2) * t) + _
        nucl.P * nucl.S3 * nucl.F3 * Exp(-(nucl.L + E * glob.rho / glob.L3) * t) _
        )
End Function
Private Sub getJacobian(ByVal opt As String, nuclide1 As MyNuclide, nuclide2 As MyNuclide, ByRef a As Double, ByRef B As Double, ByRef c As Double, ByRef d As Double, ByVal x As Double, ByVal y As Double)
    If (opt = "age_erosion") Then
        a = dNdE(nuclide1, x, y, 0)
        c = dNdE(nuclide2, x, y, 0)
        B = dNdt(nuclide1, x, y, 0)
        d = dNdt(nuclide2, x, y, 0)
    ElseIf (opt = "burial_erosion") Then
        a = dNdE(nuclide1, x, "inf", y)
        c = dNdE(nuclide2, x, "inf", y)
        B = dNdTau(nuclide1, x, "inf", y)
        d = dNdTau(nuclide2, x, "inf", y)
    ElseIf (opt = "burial_exposure") Then
        a = dNdt(nuclide1, 0, x, y)
        c = dNdt(nuclide2, 0, x, y)
        B = dNdTau(nuclide1, 0, x, y)
        d = dNdTau(nuclide2, 0, x, y)
    End If
End Sub
Private Function fBe(ByVal N1 As Double, ByVal N2 As Double, ByVal E As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide) As Double
    fBe = nuclide2.L * Log( _
        nuclide1.P * nuclide1.S0 * nuclide1.F0 / (N1 * (nuclide1.L + E * glob.rho / glob.L0)) + _
        nuclide1.P * nuclide1.S1 * nuclide1.F1 / (N1 * (nuclide1.L + E * glob.rho / glob.L1)) + _
        nuclide1.P * nuclide1.S2 * nuclide1.F2 / (N1 * (nuclide1.L + E * glob.rho / glob.L2)) + _
        nuclide1.P * nuclide1.S3 * nuclide1.F3 / (N1 * (nuclide1.L + E * glob.rho / glob.L3))) - _
        nuclide1.L * Log( _
        nuclide2.P * nuclide2.S0 * nuclide2.F0 / (N2 * (nuclide2.L + E * glob.rho / glob.L0)) + _
        nuclide2.P * nuclide2.S1 * nuclide2.F1 / (N2 * (nuclide2.L + E * glob.rho / glob.L1)) + _
        nuclide2.P * nuclide2.S2 * nuclide2.F2 / (N2 * (nuclide2.L + E * glob.rho / glob.L2)) + _
        nuclide2.P * nuclide2.S3 * nuclide2.F3 / (N2 * (nuclide2.L + E * glob.rho / glob.L3)))
End Function
Private Function DfBeDe(ByVal N1 As Double, ByVal N2 As Double, ByVal E As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide) As Double
    DfBeDe = -nuclide2.L * ( _
              nuclide1.P * nuclide1.S0 * nuclide1.F0 * glob.rho / (glob.L0 * N1 * (nuclide1.L + E * glob.rho / glob.L0) ^ 2) + _
              nuclide1.P * nuclide1.S1 * nuclide1.F1 * glob.rho / (glob.L1 * N1 * (nuclide1.L + E * glob.rho / glob.L0) ^ 2) + _
              nuclide1.P * nuclide1.S2 * nuclide1.F2 * glob.rho / (glob.L2 * N1 * (nuclide1.L + E * glob.rho / glob.L0) ^ 2) + _
              nuclide1.P * nuclide1.S3 * nuclide1.F3 * glob.rho / (glob.L3 * N1 * (nuclide1.L + E * glob.rho / glob.L0) ^ 2)) / _
            (nuclide1.P * nuclide1.S0 * nuclide1.F0 / (N1 * (nuclide1.L + E * glob.rho / glob.L0)) + _
             nuclide1.P * nuclide1.S1 * nuclide1.F1 / (N1 * (nuclide1.L + E * glob.rho / glob.L1)) + _
             nuclide1.P * nuclide1.S2 * nuclide1.F2 / (N1 * (nuclide1.L + E * glob.rho / glob.L2)) + _
             nuclide1.P * nuclide1.S3 * nuclide1.F3 / (N1 * (nuclide1.L + E * glob.rho / glob.L3))) + _
             nuclide1.L * ( _
              nuclide2.P * nuclide2.S0 * nuclide2.F0 * glob.rho / (glob.L0 * N2 * (nuclide2.L + E * glob.rho / glob.L0) ^ 2) + _
              nuclide2.P * nuclide2.S1 * nuclide2.F1 * glob.rho / (glob.L1 * N2 * (nuclide2.L + E * glob.rho / glob.L0) ^ 2) + _
              nuclide2.P * nuclide2.S2 * nuclide2.F2 * glob.rho / (glob.L2 * N2 * (nuclide2.L + E * glob.rho / glob.L0) ^ 2) + _
              nuclide2.P * nuclide2.S3 * nuclide2.F3 * glob.rho / (glob.L3 * N2 * (nuclide2.L + E * glob.rho / glob.L0) ^ 2)) / _
            (nuclide2.P * nuclide2.S0 * nuclide2.F0 / (N2 * (nuclide2.L + E * glob.rho / glob.L0)) + _
             nuclide2.P * nuclide2.S1 * nuclide2.F1 / (N2 * (nuclide2.L + E * glob.rho / glob.L1)) + _
             nuclide2.P * nuclide2.S2 * nuclide2.F2 / (N2 * (nuclide2.L + E * glob.rho / glob.L2)) + _
             nuclide2.P * nuclide2.S3 * nuclide2.F3 / (N2 * (nuclide2.L + E * glob.rho / glob.L3)))
End Function
Private Function fBt(ByVal N1 As Double, ByVal N2 As Double, ByVal t As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide) As Double
    sp1 = nuclide1.P * (nuclide1.S0 * nuclide1.F0 + nuclide1.S1 * nuclide1.F1 + nuclide1.S2 * nuclide1.F2 + nuclide1.S3 * nuclide1.F3)
    sp2 = nuclide2.P * (nuclide2.S0 * nuclide2.F0 + nuclide2.S1 * nuclide2.F1 + nuclide2.S2 * nuclide2.F2 + nuclide2.S3 * nuclide2.F3)
    fBt = nuclide2.L * Log(sp1 * (1 - Exp(-nuclide1.L * t)) / (N1 * nuclide1.L)) - _
          nuclide1.L * Log(sp2 * (1 - Exp(-nuclide2.L * t)) / (N2 * nuclide2.L))
End Function
Private Function DfBtDt(ByVal N1 As Double, ByVal N2 As Double, ByVal t As Double, nuclide1 As MyNuclide, nuclide2 As MyNuclide) As Double
    sp1 = nuclide1.P * (nuclide1.S0 * nuclide1.F0 + nuclide1.S1 * nuclide1.F1 + nuclide1.S2 * nuclide1.F2 + nuclide1.S3 * nuclide1.F3)
    sp2 = nuclide2.P * (nuclide2.S0 * nuclide2.F0 + nuclide2.S1 * nuclide2.F1 + nuclide2.S2 * nuclide2.F2 + nuclide2.S3 * nuclide2.F3)
    DfBtDt = nuclide2.L * nuclide1.L * Exp(-nuclide1.L * t) / (1 - Exp(-nuclide1.L * t)) - _
             nuclide1.L * nuclide2.L * Exp(-nuclide2.L * t) / (1 - Exp(-nuclide2.L * t))
End Function
