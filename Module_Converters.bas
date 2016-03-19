Attribute VB_Name = "Module_Converters"
Public Sub zToP(theRange As MyRange)
' convert elevation to pressure
    Dim z As Variant, P As Double
    Call addTitle(theRange, "Pressure (mbar)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        z = theRange.CellValue(rownum, 1)
        If z <> "" And IsNumeric(z) Then
            P = z2p(z)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = P
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = P
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub zToAntP(theRange As MyRange)
' convert elevation to Antarctic pressure
    Dim z As Variant, antP As Double
    Call addTitle(theRange, "Antarctic Pressure (mbar)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        z = theRange.CellValue(rownum, 1)
        If z <> "" And IsNumeric(z) Then
            antP = z2antP(z)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = antP
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = antP
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub zToD(theRange As MyRange)
' convert elevation to atmospheric depth
    Dim z As Variant, P As Double, d As Double
    Call addTitle(theRange, "Atmospheric depth (g/cm2)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        z = theRange.CellValue(rownum, 1)
        If z <> "" And IsNumeric(z) Then
            P = z2p(z)
            d = p2d(P)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = d
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = d
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub pToZ(theRange As MyRange)
' convert pressure to elevation
    Dim P As Variant, z As Double
    Call addTitle(theRange, "Elevation (m)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        P = theRange.CellValue(rownum, 1)
        If P <> "" And IsNumeric(P) Then
            z = p2z(P)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = z
                .NumberFormat = "0"
            Else
                .Offset(0, 1) = z
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub pToAntP(theRange As MyRange)
' convert standard atmospheric pressure to Antarctic pressure
    Dim z As Double, P As Variant, antP As Double
    Call addTitle(theRange, "Antarctic Pressure (mbar)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        P = theRange.CellValue(rownum, 1)
        If P <> "" And IsNumeric(P) Then
            z = p2z(P)
            antP = z2antP(z)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = antP
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = antP
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub pToD(theRange As MyRange)
' convert pressure to atmospheric depth
    Dim P As Variant, d As Double
    Call addTitle(theRange, "Atmospheric depth (g/cm2)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        P = theRange.CellValue(rownum, 1)
        If P <> "" And IsNumeric(P) Then
            d = p2d(P)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = d
                .Offset(0, 1).NumberFormat = "0"
            Else
                .Offset(0, 1).Value = d
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub AntPtoZ(theRange As MyRange)
' convert Antarctic pressure to elevation
    Dim antP As Variant, z As Double
    Call addTitle(theRange, "Elevation (m)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        antP = theRange.CellValue(rownum, 1)
        If antP <> "" And IsNumeric(antP) Then
            z = antP2z(antP)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = z
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = z
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub AntPtoP(theRange As MyRange)
' convert Antarctic pressure to standard atmospheric pressure
    Dim antP As Variant, z As Double, P As Double
    Call addTitle(theRange, "Pressure (mbar)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        antP = theRange.CellValue(rownum, 1)
        If antP <> "" And IsNumeric(antP) Then
            z = antP2z(antP)
            P = z2p(z)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = P
                .Offset(0, 1).NumberFormat = "0"
            Else
                .Offset(0, 1).Value = P
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub AntPtoD(theRange As MyRange)
' convert Antarctic pressure to standard atmospheric depth
    Dim antP As Variant, z As Double, P As Double, d As Double
    Call addTitle(theRange, "Atmospheric depth (g/cm2)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        antP = theRange.CellValue(rownum, 1)
        If antP <> "" And IsNumeric(antP) Then
            z = antP2z(antP)
            P = z2p(z)
            d = p2d(P)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = d
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = d
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub dToZ(theRange As MyRange)
' convert atmospheric depth to elevation
    Dim d As Variant, P As Double, z As Double
    Call addTitle(theRange, "Elevation (m)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        d = theRange.CellValue(rownum, 1)
        If d <> "" And IsNumeric(d) Then
            P = d2p(d)
            z = p2z(P)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = z
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = z
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub dToP(theRange As MyRange)
' convert atmospheric depth to pressure
    Dim d As Variant, P As Double
    Call addTitle(theRange, "Pressure (mbar)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        d = theRange.CellValue(rownum, 1)
        If d <> "" And IsNumeric(d) Then
            P = d2p(d)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = P
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = P
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub dToAntP(theRange As MyRange)
' convert atmospheric depth to Antarctic pressure
    Dim d As Variant, P As Double, z As Double, antP As Double
    Call addTitle(theRange, "Antarctic Pressure (mbar)")
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        d = theRange.CellValue(rownum, 1)
        If d <> "" And IsNumeric(d) Then
            P = d2p(d)
            z = p2z(P)
            antP = z2antP(z)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = antP
                .NumberFormat = "0"
            Else
                .Offset(0, 1).Value = antP
                .Offset(0, 1).NumberFormat = "0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub degToDegMin(theRange As MyRange)
    Dim deg As Variant, min As Double
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, 1))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = "deg"
            .Offset(-1, 1).Value = "min"
        Else
            .Offset(-1, 1).Value = "deg"
            .Offset(-1, 2).Value = "min"
        End If
    End If
    End With
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            min = 60 * (deg - floor(deg))
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = floor(deg)
                .NumberFormat = "0"
                .Offset(0, 1).Value = min
                .Offset(0, 1).NumberFormat = "0.00"
            Else
                .Offset(0, 1).Value = floor(deg)
                .Offset(0, 1).NumberFormat = "0"
                .Offset(0, 2).Value = min
                .Offset(0, 2).NumberFormat = "0.00"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub degToDegMinSec(theRange As MyRange)
    Dim deg As Variant, min As Double, sec As Double
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, 1))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = "deg"
            .Offset(-1, 1).Value = "min"
            .Offset(-1, 2).Value = "sec"
        Else
            .Offset(-1, 1).Value = "deg"
            .Offset(-1, 2).Value = "min"
            .Offset(-1, 3).Value = "sec"
        End If
    End If
    End With
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            min = 60 * (deg - floor(deg))
            sec = 60 * (min - floor(min))
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = floor(deg)
                .NumberFormat = "0"
                .Offset(0, 1).Value = floor(min)
                .Offset(0, 1).NumberFormat = "0"
                .Offset(0, 2).Value = sec
                .Offset(0, 2).NumberFormat = "0.0"
            Else
                .Offset(0, 1).Value = floor(deg)
                .Offset(0, 1).NumberFormat = "0"
                .Offset(0, 2).Value = floor(min)
                .Offset(0, 2).NumberFormat = "0"
                .Offset(0, 3).Value = sec
                .Offset(0, 3).NumberFormat = "0.0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub degToI(theRange As MyRange)
    Dim deg As Variant, incl As Double
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, 1))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = "Inclination"
        Else
            .Offset(-1, 1).Value = "Inclination"
        End If
    End If
    End With
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            incl = deg2i(deg)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = incl
                .NumberFormat = "0.0"
            Else
                .Offset(0, 1).Value = incl
                .Offset(0, 1).NumberFormat = "0.0"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub iToDeg(theRange As MyRange)
    Dim incl As Variant, deg As Double
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, 1))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = "deg"
        Else
            .Offset(-1, 1).Value = "deg"
        End If
    End If
    End With
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        incl = theRange.CellValue(rownum, 1)
        If incl <> "" And IsNumeric(incl) Then
            deg = i2deg(incl)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = deg
                .NumberFormat = "0.00"
            Else
                .Offset(0, 1).Value = deg
                .Offset(0, 1).NumberFormat = "0.00"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub RcToDeg(theRange As MyRange)
    Dim Rc As Variant, deg As Double
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, 1))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = "deg"
        Else
            .Offset(-1, 1).Value = "deg"
        End If
    End If
    End With
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        Rc = theRange.CellValue(rownum, 1)
        If Rc <> "" And IsNumeric(Rc) Then
            If Rc <= 0 Then
                Rc = glob.Zero ' Rc = 0 and less blows up
            End If
            deg = r2d(Rc)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = deg
                .NumberFormat = "0.00"
            Else
                .Offset(0, 1).Value = deg
                .Offset(0, 1).NumberFormat = "0.00"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub degToRc(theRange As MyRange)
    Dim deg As Variant, Rc As Double
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, 1))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = "Rc (GV)"
        Else
            .Offset(-1, 1).Value = "Rc (GV)"
        End If
    End If
    End With
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            If deg > 55 Then
                deg = 55 ' above 55 degrees, the polynomial fit doesn't work and doesn't matter
            End If
            Rc = d2r(deg)
            With Worksheets(theRange.sheet).Range(theRange.CellAddress(rownum, 1))
            If glob.Replace Then
                .Value = Rc
                .NumberFormat = "0.00"
            Else
                .Offset(0, 1).Value = Rc
                .Offset(0, 1).NumberFormat = "0.00"
            End If
            End With
        End If
    Next rownum
End Sub
Public Sub degMinToDeg(theRange As MyRange)
    Dim deg As Variant, min As Double
    If theRange.numcols = 2 Then
        With Worksheets(theRange.sheet)
        If .Range(theRange.CellAddress(1, 1)).row > 1 Then
            If glob.Replace Then
                .Range(theRange.CellAddress(1, 1)).Offset(-1, 0).Value = "deg"
                .Range(theRange.CellAddress(1, 2)).Offset(-1, 0).Value = ""
            Else
                .Range(theRange.CellAddress(1, 1)).Offset(-1, 2).Value = "deg"
            End If
        End If
        For rownum = 1 To theRange.numRows
            On Error Resume Next
            deg = theRange.CellValue(rownum, 1)
            If deg <> "" And IsNumeric(deg) Then
                min = theRange.CellValue(rownum, 2)
                If glob.Replace Then
                    .Range(theRange.CellAddress(rownum, 1)).Value = deg + min / 60
                    .Range(theRange.CellAddress(rownum, 2)).NumberFormat = "0.00"
                    .Range(theRange.CellAddress(rownum, 2)).Value = ""
                Else
                    .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = deg + min / 60
                    .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0.00"
                End If
            End If
        Next rownum
        End With
    Else
        MsgBox ("Please select two columns of data")
    End If
End Sub
Public Sub degMinToDegMinSec(theRange As MyRange)
Dim deg As Variant, min As Double, sec As Double
If theRange.numcols = 2 Then
    With Worksheets(theRange.sheet)
    If .Range(theRange.CellAddress(1, 1)).row > 1 Then
        If glob.Replace Then
            .Range(theRange.CellAddress(1, 1)).Offset(-1, 0).Value = "deg"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 0).Value = "min"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 1).Value = "sec"
        Else
            .Range(theRange.CellAddress(1, 1)).Offset(-1, 2).Value = "deg"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 2).Value = "min"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 3).Value = "sec"
        End If
    End If
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            min = theRange.CellValue(rownum, 2)
            sec = 60 * (min - floor(min))
            If glob.Replace Then
                    .Range(theRange.CellAddress(rownum, 1)).Value = floor(deg)
                    .Range(theRange.CellAddress(rownum, 1)).NumberFormat = "0"
                    .Range(theRange.CellAddress(rownum, 2)).Value = floor(min)
                    .Range(theRange.CellAddress(rownum, 2)).NumberFormat = "0"
                    .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = sec
                    .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0.0"
            Else
                .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).Value = floor(deg)
                .Range(theRange.CellAddress(rownum, 2)).Offset(0, 1).NumberFormat = "0"
                .Range(theRange.CellAddress(rownum, 2)).Offset(0, 2).Value = floor(min)
                .Range(theRange.CellAddress(rownum, 2)).Offset(0, 2).NumberFormat = "0"
                .Range(theRange.CellAddress(rownum, 2)).Offset(0, 3).Value = sec
                .Range(theRange.CellAddress(rownum, 2)).Offset(0, 3).NumberFormat = "0.0"
            End If
        End If
    Next rownum
    End With
Else
    MsgBox ("Please select two columns")
End If
End Sub
Public Sub degMinSecToDeg(theRange As MyRange)
Dim deg As Variant, min As Double, sec As Double
If theRange.numcols = 3 Then
    With Worksheets(theRange.sheet)
    If Range(theRange.CellAddress(1, 1)).row > 1 Then
        If glob.Replace Then
            .Range(theRange.CellAddress(1, 1)).Offset(-1, 0).Value = "deg"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 0).Value = ""
            .Range(theRange.CellAddress(1, 3)).Offset(-1, 0).Value = ""
        Else
            .Range(theRange.CellAddress(1, 3)).Offset(-1, 1).Value = "deg"
        End If
    End If
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            min = theRange.CellValue(rownum, 2)
            sec = theRange.CellValue(rownum, 3)
            If glob.Replace Then
                .Range(theRange.CellAddress(rownum, 1)).Value = deg + min / 60 + sec / 3600
                .Range(theRange.CellAddress(rownum, 1)).NumberFormat = "0.00"
                .Range(theRange.CellAddress(rownum, 2)).Value = ""
                .Range(theRange.CellAddress(rownum, 3)).Value = ""
            Else
                .Range(theRange.CellAddress(rownum, 3)).Offset(0, 1).Value = deg + min / 60 + sec / 3600
                .Range(theRange.CellAddress(rownum, 3)).Offset(0, 1).NumberFormat = "0.00"
            End If
        End If
    Next rownum
    End With
Else
    MsgBox ("Please select three columns")
End If
End Sub
Public Sub degMinSecToDegMin(theRange As MyRange)
Dim deg As Variant, min As Double, sec As Double
If theRange.numcols = 3 Then
    With Worksheets(theRange.sheet)
    If Range(theRange.CellAddress(1, 1)).row > 1 Then
        If glob.Replace Then
            .Range(theRange.CellAddress(1, 1)).Offset(-1, 0).Value = "deg"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 0).Value = "min"
            .Range(theRange.CellAddress(1, 3)).Offset(-1, 0).Value = ""
        Else
            .Range(theRange.CellAddress(1, 1)).Offset(-1, 3).Value = "deg"
            .Range(theRange.CellAddress(1, 2)).Offset(-1, 3).Value = "min"
        End If
    End If
    For rownum = 1 To theRange.numRows
        On Error Resume Next
        deg = theRange.CellValue(rownum, 1)
        If deg <> "" And IsNumeric(deg) Then
            min = theRange.CellValue(rownum, 2)
            sec = theRange.CellValue(rownum, 3)
            min = min + sec / 60
            If glob.Replace Then
                .Range(theRange.CellAddress(rownum, 1)).Value = floor(deg)
                .Range(theRange.CellAddress(rownum, 1)).NumberFormat = "0"
                .Range(theRange.CellAddress(rownum, 2)).Value = min
                .Range(theRange.CellAddress(rownum, 2)).NumberFormat = "0.00"
                .Range(theRange.CellAddress(rownum, 3)).Value = ""
            Else
                .Range(theRange.CellAddress(rownum, 3)).Offset(0, 1).Value = floor(deg)
                .Range(theRange.CellAddress(rownum, 3)).Offset(0, 1).NumberFormat = "0"
                .Range(theRange.CellAddress(rownum, 3)).Offset(0, 2).Value = min
                .Range(theRange.CellAddress(rownum, 3)).Offset(0, 2).NumberFormat = "0.00"
            End If
        End If
    Next rownum
    End With
Else
    MsgBox ("Please select three columns")
End If
End Sub
Private Function z2p(ByVal z As Double) As Double
    z2p = glob.P_o * (1 - (glob.B_o * z / glob.T_o)) ^ (glob.G_o / (glob.R_d * glob.B_o))
End Function
Private Function p2z(ByVal P As Double) As Double
    p2z = (glob.T_o / glob.B_o) * (1 - (P / glob.P_o) ^ ((glob.R_d * glob.B_o) / glob.G_o))
End Function
Private Function z2antP(ByVal z As Double) As Double
    z2antP = 989.1 * Exp(-z / 7588)
End Function
Private Function p2d(ByVal P As Double) As Double
    p2d = 10 * P / glob.G_o
End Function
Private Function antP2z(ByVal antP As Double) As Double
    antP2z = -7588 * Log(antP / 989.1)
End Function
Private Function d2p(ByVal d As Double) As Double
    d2p = d * glob.G_o / 10
End Function
Private Function deg2i(ByVal deg As Double) As Double
    deg2i = Abs(Atn(2 * Tan(pi * deg / 180)) * 180 / pi)
    If deg < 0 Then
        deg2i = -deg2i
    End If
End Function
Private Function i2deg(ByVal incl As Double) As Double
    i2deg = Atn(0.5 * Tan(pi * Abs(incl) / 180)) * 180 / pi
    If incl < 0 Then
        i2deg = -i2deg
    End If
End Function
Private Function r2d(ByVal Rc As Double) As Double
' convert cut-off rigidity to degrees latitude
    deg = 45 'initial guess
    ' find root with Newton's method
    For i = 1 To 100
        RcEst = d2r(deg)
        If Abs((RcEst - Rc) / Rc) < glob.Zero Then
            Exit For
        End If
        e0 = -0.0043077
        E1 = 0.024352
        E2 = -0.0046757
        E3 = 0.00033287
        e4 = -0.000010993
        e5 = 0.00000017037
        e6 = -0.0000000010043
        F0 = 14.792
        F1 = -0.066799
        F2 = 0.0035714
        F3 = 0.000028005
        f4 = -0.000023902
        f5 = 0.00000066179
        f6 = -0.0000000050283
        dRdDeg = (E1 + F1 * glob.MM0) + (E2 + F2 * glob.MM0) * 2 * deg + _
                 (E3 + F3 * glob.MM0) * 3 * deg ^ 2 + (e4 + f4 * glob.MM0) * 4 * deg ^ 3 + _
                 (e5 + f5 * glob.MM0) * 5 * deg ^ 4 + (e6 + f6 * glob.MM0) * 6 * deg ^ 5
        dDeg = (RcEst - Rc) / dRdDeg
        deg = deg - dDeg
    Next i
    r2d = deg
End Function
Private Function d2r(ByVal deg As Double) As Double
' convert cut-off degrees latitude to rigidity
    d = Abs(deg)
    e0 = -0.0043077
    E1 = 0.024352
    E2 = -0.0046757
    E3 = 0.00033287
    e4 = -0.000010993
    e5 = 0.00000017037
    e6 = -0.0000000010043
    F0 = 14.792
    F1 = -0.066799
    F2 = 0.0035714
    F3 = 0.000028005
    f4 = -0.000023902
    f5 = 0.00000066179
    f6 = -0.0000000050283
    d2r = (e0 + F0 * glob.MM0) + _
         (E1 + F1 * glob.MM0) * d + (E2 + F2 * glob.MM0) * d ^ 2 + _
         (E3 + F3 * glob.MM0) * d ^ 3 + (e4 + f4 * glob.MM0) * d ^ 4 + _
         (e5 + f5 * glob.MM0) * d ^ 5 + (e6 + f6 * glob.MM0) * d ^ 6
End Function
Private Sub addTitle(theRange As MyRange, title As String)
    With Worksheets(theRange.sheet).Range(theRange.CellAddress(1, theRange.numcols))
    If .row > 1 Then
        If glob.Replace Then
            .Offset(-1, 0).Value = title
        Else
            .Offset(-1, 1).Value = title
        End If
    End If
    End With
End Sub
Private Function floor(ByVal num As Double) As Integer
    floor = WorksheetFunction.Round(num, 0)
    If num - floor < 0 Then
        floor = floor - 1
    End If
End Function
