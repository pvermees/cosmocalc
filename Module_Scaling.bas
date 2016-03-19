Attribute VB_Name = "Module_Scaling"
Public Sub LalScaling(theRange As MyRange)
    Dim latitude As Variant
    Dim elevation As Double
    Dim row As Range

    With Worksheets(theRange.sheet)
    If theRange.numcols = 2 Then
        If .Range(theRange.CellAddress(1, 2)).row > 1 Then
           .Range(theRange.CellAddress(1, 2)).Offset(-1, 1).Value = "Lal scaling"
        End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        latitude = theRange.CellValue(rownum, 1)
        If latitude <> "" And IsNumeric(latitude) Then
            elevation = theRange.CellValue(rownum, 2) / 1000
            S = LalFact(latitude, elevation)
Label:
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = S
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0.00"
errorhandler:
            If Err.Number <> 0 Then
                S = 0
                Resume Label
            End If
        End If
    Next rownum
    Else
        MsgBox ("Please select two columns of data")
    End If
    End With
End Sub
Public Sub StoneScaling(theRange As MyRange, nucl As MyNuclide)
    Dim latitude As Variant
    Dim row As Range

    With Worksheets(theRange.sheet)
    If theRange.numcols = 2 Then
        If .Range(theRange.CellAddress(1, 2)).row > 1 Then
           .Range(theRange.CellAddress(1, 2)).Offset(-1, 1).Value = "Stone scaling"
        End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        latitude = theRange.CellValue(rownum, 1)
        If latitude <> "" And IsNumeric(latitude) Then
            pressure = theRange.CellValue(rownum, 2)
            S = StoneFact(latitude, pressure, nucl)
Label:
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = S
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0.00"
errorhandler:
            If Err.Number <> 0 Then
                S = 0
                Resume Label
            End If
        End If
    Next rownum
    Else
        MsgBox ("Please select two columns of data")
    End If
    End With
End Sub
Public Sub DunaiScaling(theRange As MyRange, nucl As MyNuclide)
    Dim depth As Double
    Dim inclination As Variant
    
    With Worksheets(theRange.sheet)
    If theRange.numcols = 2 Then
        If .Range(theRange.CellAddress(1, 2)).row > 1 Then
           .Range(theRange.CellAddress(1, 2)).Offset(-1, 1).Value = "Dunai scaling"
        End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        inclination = theRange.CellValue(rownum, 1)
        If inclination <> "" And IsNumeric(inclination) Then
            depth = theRange.CellValue(rownum, 2)
            S = DunaiFact(inclination, depth, nucl)
Label:
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = S
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0.00"
errorhandler:
            If Err.Number <> 0 Then
                S = 0
                Resume Label
            End If
        End If
    Next rownum
    Else
        MsgBox ("Please select two columns of data")
    End If
    End With
End Sub
Public Sub DesiletScaling(ByVal year As Integer, theRange As MyRange, nucl As MyNuclide)
    Dim depth As Double
    Dim Rc2 As Variant
    
    With Worksheets(theRange.sheet)
    If theRange.numcols = 2 Then
        If .Range(theRange.CellAddress(1, 2)).row > 1 Then
           .Range(theRange.CellAddress(1, 2)).Offset(-1, 1).Value = "Desilets scaling"
        End If
    For rownum = 1 To theRange.numRows
        On Error GoTo errorhandler
        Rc2 = theRange.CellValue(rownum, 1)
        If Rc2 <> "" And IsNumeric(Rc2) Then
            If Rc2 = 0 Then
                Rc2 = 1
            End If
            depth2 = theRange.CellValue(rownum, 2)
            S = DesiletsFact(Rc2, depth2, year, nucl)
Label:
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = S
            .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0.00"
errorhandler:
            If Err.Number <> 0 Then
                S = 0
                Resume Label
            End If
        End If
    Next rownum
    Else
        MsgBox ("Please select two columns of data")
    End If
    End With
End Sub
Private Function LalFact(ByVal mylat As Double, elevation As Double) As Double
    Dim a1i As Double, a2i As Double, a3i As Double, a4i As Double, drate1, drate2 As Double
    Dim Lat(10) As Double
    Dim a1(10) As Double
    Dim a2(10) As Double
    Dim a3(10) As Double
    Dim a4(10) As Double
    
    latitude = Abs(mylat)
    
    Lat(1) = 0
    Lat(2) = 10
    Lat(3) = 20
    Lat(4) = 30
    Lat(5) = 40
    Lat(6) = 50
    Lat(7) = 60
    Lat(8) = 70
    Lat(9) = 80
    Lat(10) = 90
    
    a1(1) = 330.7
    a1(2) = 337.9
    a1(3) = 382.1
    a1(4) = 469.3
    a1(5) = 525.6
    a1(6) = 571.1
    a1(7) = 563.4
    a1(8) = 563.4
    a1(9) = 563.4
    a1(10) = 563.4
    
    a2(1) = 255.9
    a2(2) = 252.1
    a2(3) = 272.1
    a2(4) = 394.6
    a2(5) = 505.4
    a2(6) = 588.1
    a2(7) = 621.8
    a2(8) = 621.8
    a2(9) = 621.8
    a2(10) = 621.8
    
    a3(1) = 98.43
    a3(2) = 111
    a3(3) = 132.5
    a3(4) = 97.76
    a3(5) = 142
    a3(6) = 170.9
    a3(7) = 177.3
    a3(8) = 177.3
    a3(9) = 177.3
    a3(10) = 177.3
    
    a4(1) = 20.5
    a4(2) = 20.73
    a4(3) = 24.83
    a4(4) = 47.2
    a4(5) = 58.87
    a4(6) = 76.12
    a4(7) = 78.91
    a4(8) = 78.91
    a4(9) = 78.91
    a4(10) = 78.91
    
    a1i = interp(latitude, Lat, a1)
    a2i = interp(latitude, Lat, a2)
    a3i = interp(latitude, Lat, a3)
    a4i = interp(latitude, Lat, a4)
    drate1 = a1i + a2i * elevation + a3i * elevation ^ 2 + a4i * elevation ^ 3
    drate2 = a1(10)
    LalFact = drate1 / drate2
End Function
Private Function StoneFact(ByVal mylat As Double, ByVal pressure As Double, nucl As MyNuclide) As Double
    Dim drate1 As Double, a1i As Double, a2i As Double, a3i As Double, a4i As Double, a5i As Double, Mi As Double, drate2, Ml1, Ml2 As Double
    Dim Lat(10) As Double
    Dim a1(10) As Double
    Dim a2(10) As Double
    Dim a3(10) As Double
    Dim a4(10) As Double
    Dim a5(10) As Double
    Dim M(10) As Double
    
    latitude = Abs(mylat)
    
    Lat(1) = 0
    Lat(2) = 10
    Lat(3) = 20
    Lat(4) = 30
    Lat(5) = 40
    Lat(6) = 50
    Lat(7) = 60
    Lat(8) = 70
    Lat(9) = 80
    Lat(10) = 90

    a1(1) = 31.8518
    a1(2) = 34.3699
    a1(3) = 40.3153
    a1(4) = 42.0983
    a1(5) = 56.7733
    a1(6) = 69.072
    a1(7) = 71.8733
    a1(8) = 71.8733
    a1(9) = 71.8733
    a1(10) = 71.8733
    
    a2(1) = 250.3193
    a2(2) = 258.4759
    a2(3) = 308.9894
    a2(4) = 512.6857
    a2(5) = 649.1343
    a2(6) = 832.4566
    a2(7) = 863.1927
    a2(8) = 863.1927
    a2(9) = 863.1927
    a2(10) = 863.1927
    
    a3(1) = -0.083393
    a3(2) = -0.089807
    a3(3) = -0.106248
    a3(4) = -0.120551
    a3(5) = -0.160859
    a3(6) = -0.199252
    a3(7) = -0.207069
    a3(8) = -0.207069
    a3(9) = -0.207069
    a3(10) = -0.207069
    
    a4(1) = 0.00007426
    a4(2) = 0.000079457
    a4(3) = 0.000094508
    a4(4) = 0.00011752
    a4(5) = 0.00015463
    a4(6) = 0.00019391
    a4(7) = 0.00020127
    a4(8) = 0.00020127
    a4(9) = 0.00020127
    a4(10) = 0.00020127
    
    a5(1) = -0.000000022397
    a5(2) = -0.000000023697
    a5(3) = -0.000000028234
    a5(4) = -0.000000038809
    a5(5) = -0.00000005033
    a5(6) = -0.000000063653
    a5(7) = -0.000000066043
    a5(8) = -0.000000066043
    a5(9) = -0.000000066043
    a5(10) = -0.000000066043
    
    M(1) = 0.587
    M(2) = 0.6
    M(3) = 0.678
    M(4) = 0.833
    M(5) = 0.933
    M(6) = 1
    M(7) = 1
    M(8) = 1
    M(9) = 1
    M(10) = 1
    
    a1i = interp(latitude, Lat, a1)
    a2i = interp(latitude, Lat, a2)
    a3i = interp(latitude, Lat, a3)
    a4i = interp(latitude, Lat, a4)
    a5i = interp(latitude, Lat, a5)
    Mi = interp(latitude, Lat, M)
    drate1 = a1i + a2i * Exp(-pressure / 150) + a3i * pressure + a4i * pressure ^ 2 + a5i * pressure ^ 3
    Ml1 = Mi * Exp((1013.25 - pressure) / 242)
    drate1 = nucl.F0 * drate1 + (1 - nucl.F0) * Ml1
    drate2 = a1(10) + a2(10) * Exp(-1013.25 / 150) + a3(10) * 1013.25 + a4(10) * 1013.25 ^ 2 + a5(10) * 1013.25 ^ 3
    Ml2 = M(10)
    drate2 = nucl.F0 * drate2 + (1 - nucl.F0) * Ml2
    StoneFact = drate1 / drate2
End Function
Private Function DunaiFact(ByVal myincl As Double, ByVal depth As Double, nucl As MyNuclide) As Double
    Dim a1 As Double, B1 As Double, C1 As Double, x1 As Double, y1 As Double, _
        a2 As Double, B2 As Double, C2 As Double, x2 As Double, y2 As Double, L1 As Double, _
        L2  As Double, N1 As Double, N2 As Double, z1 As Double, z2 As Double, _
        N1030_1 As Double, N1030_2 As Double
        
    inclination = Abs(myincl)
        
    a1 = 0.445
    B1 = 4.1703
    C1 = 0.335
    x1 = 62.698
    y1 = 0.5555
    a2 = 19.85
    B2 = -5.43
    C2 = 3.59
    x2 = 62.05
    y2 = 129.55
    
    z1 = 1033.2 - depth
    z2 = 0
    N1030_1 = y1 + a1 / ((1 + Exp(-((inclination - x1) / B1))) ^ C1)
    N1030_2 = y1 + a1 / ((1 + Exp(-((90 - x1) / B1))) ^ C1)
    L1 = y2 + a2 / ((1 + Exp(-((inclination - x2) / B2))) ^ C2)
    L2 = y2 + a2 / ((1 + Exp(-((90 - x2) / B2))) ^ C2)
    N1 = N1030_1 * Exp(z1 / L1) * nucl.F0 + N1030_1 * Exp(z1 / 247) * (1 - nucl.F0)
    N2 = N1030_2 * Exp(z2 / L2) * nucl.F0 + N1030_2 * Exp(z2 / 247) * (1 - nucl.F0)
    DunaiFact = N1 / N2
End Function
Private Function DesiletsFact(ByVal Rc2 As Double, ByVal depth2 As Double, ByVal year As Integer, nucl As MyNuclide) As Double
    Dim P1 As Double, depth1 As Double, Rc1 As Double, P2sp As Double, _
        P2muF As Double, P2muS As Double, P2 As Double

    P1 = 1
    depth1 = 1033.2
    Rc1 = 1 ' several calculations blow up when Rc1 = 0
                
    P2sp = getPsp(nucl.F0 * P1, Rc1, depth1, Rc2, depth2, year)
    P2muF = getPmuF(nucl.F3 * P1, Rc1, depth1, Rc2, depth2)
    P2muS = getPmuS((nucl.F1 + nucl.F2) * P1, Rc1, depth1, Rc2, depth2)

    P2 = P2sp + P2muF + P2muS
                
    DesiletsFact = P2
End Function
Private Function getPsp(ByVal P1 As Double, ByVal Rc1 As Double, ByVal x1 As Double, ByVal Rc2 As Double, ByVal x2 As Double, ByVal year As Integer)
'  scale the spallogenic production rate P1 at cutoff rigidity Rc1 and
' atmospheric depth x1 to cutoff rigidity Rc2 and atmospheric depth x2
Dim Lsp1, P1sea, alpha, k, P2sea, Lsp2 As Double

' scale P1 to sea level at Rc1:
Lsp1 = getLsp(Rc1, x1, 1033.2, year)
P1sea = P1 * Exp((x1 - 1033.2) / Lsp1)
' scale P1sea to Rc2:
alpha = 10.275
k = 0.9615
P2sea = P1sea * (1 - Exp(-alpha * Rc2 ^ (-k))) / (1 - Exp(-alpha * Rc1 ^ (-k)))
' scale P2sea to depth2:
Lsp2 = getLsp(Rc2, x2, 1033.2, year)
getPsp = P2sea * Exp((1033.2 - x2) / Lsp2)
End Function
Private Function getLsp(ByVal Rc As Double, ByVal x1 As Double, ByVal x2 As Double, ByVal year As Integer)
' calculate the attenuation length between depth x1 and depth x2 for cutoff rigidity Rc
' using equation 10 of Desilets and Zreda (2003)
' if (x1==x2) x2 = x1-0.1; end % otherwise L = NaN
Dim N, alpha, k, b0, B1, B2, b3, b4, b5, b6, b7, b8, term1, term2 As Double

If year = 2003 Then
    N = 0.0099741
    alpha = 0.45318
    k = -0.081613
    b0 = 0.0000063813
    B1 = -0.00000062639
    B2 = -0.0000000051187
    b3 = -0.0000000071914
    b4 = 0.0000000011291
    b5 = 0.0000000000174
    b6 = 2.5816E-12
    b7 = -5.8588E-13
    b8 = -1.1268E-14
ElseIf year = 2006 Then
    N = 0.010177
    alpha = 0.10207
    k = -0.39527
    b0 = 0.0000085236
    B1 = -0.0000006367
    B2 = -0.0000000070814
    b3 = -0.0000000099182
    b4 = 0.0000000009925
    b5 = 0.000000000024925
    b6 = 3.8615E-12
    b7 = -4.8194E-13
    b8 = -1.5371E-14
End If

' Desilet's equation blows up for very small Rc
If Rc < 1 Then
    Rc = 1
End If

' otherwise term1 = term2
If x2 = 1033.2 Then
    x1 = 1000 ' just calculate the attenuation length near sea level
End If
term1 = (N * (1 + Exp(-alpha * Rc ^ (-k))) ^ (-1) * x2 + _
    0.5 * (b0 + B1 * Rc + B2 * Rc ^ 2) * x2 ^ 2 + _
    (1 / 3) * (b3 + b4 * Rc + b5 * Rc ^ 2) * x2 ^ 3 + _
    0.25 * (b6 + b7 * Rc + b8 * Rc ^ 2) * x2 ^ 4)
term2 = (N * (1 + Exp(-alpha * Rc ^ (-k))) ^ (-1) * x1 + _
    0.5 * (b0 + B1 * Rc + B2 * Rc ^ 2) * x1 ^ 2 + _
    (1 / 3) * (b3 + b4 * Rc + b5 * Rc ^ 2) * x1 ^ 3 + _
    0.25 * (b6 + b7 * Rc + b8 * Rc ^ 2) * x1 ^ 4)
getLsp = (x2 - x1) / (term1 - term2)
End Function
Private Function getPmuF(ByVal P1 As Double, ByVal Rc1 As Double, ByVal x1 As Double, ByVal Rc2 As Double, ByVal x2 As Double)
' scale the fast muogenic production rate P1 at cutoff rigidity Rc1 and
' atmospheric depth x1 to cutoff rigidity Rc2 and atmospheric depth x2
Dim a0, a1, a2, a3, LmuF1, P1sea, alpha, k, P2sea, LmuF2 As Double

a0 = 216.58
a1 = 8.783
a2 = -0.0013532
a3 = 0.37859
LmuF1 = a0 + a1 * Rc1 + x1 * (a2 * Rc1 + a3)
P1sea = P1 * Exp((x1 - 1033.2) / LmuF1)
' scale P1sea to Rc2:
alpha = 38.51
k = 1.03
P2sea = P1sea * (1 - Exp(-alpha * Rc2 ^ (-k))) / (1 - Exp(-alpha * Rc1 ^ (-k)))
' scale P2sea to depth2:
LmuF2 = a0 + a1 * Rc2 + x2 * (a2 * Rc2 + a3)
getPmuF = P2sea * Exp((1033.2 - x2) / LmuF2)
End Function
Private Function getPmuS(ByVal P1 As Double, ByVal Rc1 As Double, ByVal x1 As Double, ByVal Rc2 As Double, ByVal x2 As Double)
' scale the slow muogenic production rate P1 at cutoff rigidity Rc1 and
' atmospheric depth x1 to cutoff rigidity Rc2 and atmospheric depth x2
Dim LmuS1, P1sea, alpha, k, P2sea, LmuS2 As Double

LmuS1 = 233 + 3.68 * Rc1
P1sea = P1 * Exp((x1 - 1033.2) / LmuS1)
' scale P1sea to Rc2:
alpha = 38.51
k = 1.03
P2sea = P1sea * (1 - Exp(-alpha * Rc2 ^ (-k))) / (1 - Exp(-alpha * Rc1 ^ (-k)))
' scale P2sea to depth2:
' if (x2==1033.2) x2 = 1033.1; end % otherwise Lsp2 = NaN
LmuS2 = 233 + 3.68 * Rc2
getPmuS = P2sea * Exp((1033.2 - x2) / LmuS2)
End Function
Private Function interp(ByVal latitude As Double, Lat() As Double, Arr() As Double) As Double
    Dim i As Integer
    For i = 1 To UBound(Lat) - 1
        If latitude >= Lat(i) And latitude <= Lat(i + 1) Then
            interp = Arr(i) + ((latitude - Lat(i)) / (Lat(i + 1) - Lat(i))) * (Arr(i + 1) - Arr(i))
            Exit For
        End If
    Next i
End Function
