VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm 
   Caption         =   "Settings"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   -105
   ClientWidth     =   10680
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub userform_initialize()
    On Error Resume Next
    With Me.ComboBox1
        .AddItem "Lal"
        .AddItem "Stone"
        .AddItem "Dunai"
        .AddItem "Desilets & Zreda (2003)"
        .AddItem "Desilets et al (2006)"
    End With
    With Me.ComboBox2
        .AddItem "Braucher"
        .AddItem "Granger"
        .AddItem "Spallation only"
        .AddItem "Schaller"
        .Text = glob.equation
    End With
    Call showVals1
    Call showVals2
    Call showVals3
End Sub
Private Sub tieNe2Be_Click()
    glob.tieNe2Be = tieNe2Be.Value
    Call glob.setSLHLp("21Ne")
    Call showVals1
End Sub
Private Sub addRecordButton_Click()
    On Error Resume Next
    reminder.Visible = True
    Me.nBox.Text = ""
    Me.ageBox.Text = ""
    Me.latBox.Text = ""
    Me.elevBox.Text = ""
    Me.PBox.Text = ""
    Me.refBox.Text = ""
    Me.numrecBox = SpinButton.Max + 1
    Me.recnumBox = SpinButton.Max + 1
    End Sub
Private Sub enterRecordButton_Click()
    On Error Resume Next
    Dim N As Integer
    nucl = TabStrip.SelectedItem.Caption
    If (Me.reminder.Visible) Then
        Dim theRange As MyRange
        Set theRange = New MyRange
        If Not (Me.nBox.Text = "" Or Me.ageBox.Text = "" Or Me.latBox.Text = "" Or Me.elevBox.Text = "") _
            And (IsNumeric(CDbl(Me.nBox.Text)) And IsNumeric(CDbl(Me.ageBox.Text)) _
            And IsNumeric(CDbl(Me.latBox.Text)) And IsNumeric(CDbl(Me.elevBox.Text))) Then
            N = Me.recnumBox.Value
            If N > SpinButton.Max Then
                SpinButton.Max = SpinButton.Max + 1
                Select Case nucl
                Case Is = "10Be"
                    glob.n10BeCals = N
                Case Is = "26Al"
                    glob.n26AlCals = N
                Case Is = "21Ne"
                    glob.n21NeCals = N
                Case Is = "3He"
                    glob.n3HeCals = N
                Case Is = "36Cl"
                    glob.n36ClCals = N
                Case Is = "14C"
                    glob.n14CCals = N
                End Select
            End If
            Call theRange.SetProperties(glob.CalRange(nucl), glob.name)
            Call theRange.SetCellValue(N, 1, CDbl(Me.nBox.Text))
            Call theRange.SetCellValue(N, 2, CDbl(Me.ageBox.Text))
            Call theRange.SetCellValue(N, 3, CDbl(Me.latBox.Text))
            Call theRange.SetCellValue(N, 4, CDbl(Me.elevBox.Text))
            glob.setSLHLp (nucl)
        End If
        Call showVals1
        Set theRange = Nothing
    Else
        Select Case nucl
            Case Is = "10Be"
                glob.P10Be = CDbl(Me.avgPBox.Text)
            Case Is = "26Al"
                glob.P26Al = CDbl(Me.avgPBox.Text)
            Case Is = "21Ne"
                glob.P21Ne10Be = CDbl(Me.P21Ne10Be.Text)
                Call glob.setSLHLp("21Ne")
                Call showVals1
                glob.P21Ne = CDbl(Me.avgPBox.Text)
            Case Is = "3He"
                glob.P3He = CDbl(Me.avgPBox.Text)
            Case Is = "36Cl"
                glob.P36Cl = CDbl(Me.avgPBox.Text)
            Case Is = "14C"
                glob.P14C = CDbl(Me.avgPBox.Text)
        End Select
    End If
End Sub
Private Sub deleteRecordButton_Click()
    On Error Resume Next
    Dim N As Integer
    Dim nucl As String
    nucl = TabStrip.SelectedItem.Caption
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(glob.CalRange(nucl), glob.name)
    If SpinButton.Max = 0 Then
        ' do nothing
    Else
        N = Me.recnumBox.Value
        Worksheets(glob.name).Range(glob.RecRange(nucl, N)).Delete Shift:=xlUp
        Select Case nucl
        Case Is = "10Be"
            glob.n10BeCals = glob.n10BeCals - 1
        Case Is = "26Al"
            glob.n26AlCals = glob.n26AlCals - 1
        Case Is = "21Ne"
            glob.n21NeCals = glob.n21NeCals - 1
        Case Is = "3He"
            glob.n3HeCals = glob.n3HeCals - 1
        Case Is = "36Cl"
            glob.n36ClCals = glob.n36ClCals - 1
        Case Is = "14C"
            glob.n14CCals = glob.n14CCals - 1
        End Select
    End If
    SpinButton.Max = SpinButton.Max - 1
    glob.setSLHLp (nucl)
    Call showVals1
    Set theRange = Nothing
End Sub
Private Sub SpinButton_Change()
    On Error Resume Next
    Dim recnum As Integer
    recnum = SpinButton.Value
    Me.recnumBox.Value = recnum
    Call showRec(TabStrip.SelectedItem.Caption, recnum)
End Sub
Private Sub showRec(ByVal nucl As String, ByVal recnum As Integer)
    On Error Resume Next
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(glob.CalRange(nucl), glob.name)
    Me.nBox.Text = CStr(theRange.CellValue(recnum, 1))
    Me.ageBox.Text = CStr(theRange.CellValue(recnum, 2))
    Me.latBox.Text = CStr(theRange.CellValue(recnum, 3))
    Me.elevBox.Text = CStr(theRange.CellValue(recnum, 4))
    Me.PBox.Text = Format(theRange.CellValue(recnum, 6), "#.00")
    Me.refBox.Text = theRange.CellValue(recnum, 7)
    Set theRange = Nothing
End Sub
Private Sub TabStrip_Change()
    Call showVals1
End Sub
Private Sub CloseButton_Click()
    Unload Me
End Sub
Private Sub CancelButton2_Click()
    Call glob.setSLHLf(glob.equation)
    Unload Me
End Sub
Private Sub CancelButton3_Click()
    Unload Me
End Sub
Private Sub ComboBox1_Change()
    On Error Resume Next
    Dim oldScaling As String, newScaling As String
    oldScaling = glob.Scaling
    newScaling = Me.ComboBox1.Value
    If oldScaling <> newScaling Then
        glob.Scaling = newScaling
        Call glob.ConvertAllP(oldScaling, newScaling) ' recalculate P
        Select Case newScaling
        Case Is = "Lal"
            latLabel.Caption = "Latitude (deg)"
            elevLabel.Caption = "Elevation (m)"
        Case Is = "Stone"
            latLabel.Caption = "Latitude (deg)"
            elevLabel.Caption = "Pressure (mbar)"
        Case Is = "Dunai"
            latLabel.Caption = "Inclination (deg)"
            elevLabel.Caption = "Depth (g/cm2)"
        Case Is = "Desilets & Zreda (2003)"
            latLabel.Caption = "Cut-off rigidity (GV)"
            elevLabel.Caption = "Depth (g/cm2)"
        Case Is = "Desilets et al (2006)"
            latLabel.Caption = "Cut-off rigidity (GV)"
            elevLabel.Caption = "Depth (g/cm2)"
        End Select
        Call showVals1
    End If
End Sub
Private Sub ComboBox2_Change()
    If (ComboBox2.Value) <> glob.equation Then
        Call glob.setSLHLf(ComboBox2.Value)
    End If
    Call showVals2
End Sub
Private Sub ResetButton1_Click()
    Call glob.reset1
    Call showVals1
End Sub
Private Sub ResetButton2_Click()
    Call glob.reset2
    Call showVals2
End Sub
Private Sub ResetButton3_Click()
    Call glob.reset3
    Call showVals3
End Sub
Private Sub OKButton1_Click()
    glob.Scaling = ComboBox1.Value
    Unload Me
End Sub
Private Sub OKButton2_Click()
    On Error Resume Next
    glob.F26Al0 = CDbl(Me.F26Al0Box.Text)
    glob.F26Al1 = CDbl(Me.F26Al1Box.Text)
    glob.F26Al2 = CDbl(Me.F26Al2Box.Text)
    glob.F26Al3 = CDbl(Me.F26Al3Box.Text)
    glob.F10Be0 = CDbl(Me.F10Be0Box.Text)
    glob.F10Be1 = CDbl(Me.F10Be1Box.Text)
    glob.F10Be2 = CDbl(Me.F10Be2Box.Text)
    glob.F10Be3 = CDbl(Me.F10Be3Box.Text)
    glob.F21Ne0 = CDbl(Me.F21Ne0Box.Text)
    glob.F21Ne1 = CDbl(Me.F21Ne1Box.Text)
    glob.F21Ne2 = CDbl(Me.F21Ne2Box.Text)
    glob.F21Ne3 = CDbl(Me.F21Ne3Box.Text)
    glob.F36Cl0 = CDbl(Me.F36Cl0Box.Text)
    glob.F36Cl1 = CDbl(Me.F36Cl1Box.Text)
    glob.F36Cl2 = CDbl(Me.F36Cl2Box.Text)
    glob.F36Cl3 = CDbl(Me.F36Cl3Box.Text)
    glob.F3He0 = CDbl(Me.F3He0Box.Text)
    glob.F3He1 = CDbl(Me.F3He1Box.Text)
    glob.F3He2 = CDbl(Me.F3He2Box.Text)
    glob.F3He3 = CDbl(Me.F3He3Box.Text)
    glob.F14C0 = CDbl(Me.F14C0Box.Text)
    glob.F14C1 = CDbl(Me.F14C1Box.Text)
    glob.F14C2 = CDbl(Me.F14C2Box.Text)
    glob.F14C3 = CDbl(Me.F14C3Box.Text)
    glob.L26Al = CDbl(Me.L26AlBox.Text)
    glob.L10Be = CDbl(Me.L10BeBox.Text)
    glob.L36Cl = CDbl(Me.L36ClBox.Text)
    glob.L14C = CDbl(Me.L14CBox.Text)
    glob.L0 = CDbl(Me.L0Box.Text)
    glob.L1 = CDbl(Me.L1Box.Text)
    glob.L2 = CDbl(Me.L2Box.Text)
    glob.L3 = CDbl(Me.L3Box.Text)
    glob.rho = CDbl(Me.RhoBox.Text)
    glob.equation = Me.ComboBox2.Value
    Unload Me
End Sub
Private Sub OKButton3_Click()
    On Error Resume Next
    glob.T_o = CDbl(Me.ToBox.Text)
    glob.B_o = CDbl(Me.BoBox.Text)
    glob.P_o = CDbl(Me.PoBox.Text)
    glob.MM0 = CDbl(Me.MM0Box.Text)
    glob.exponent = CDbl(Me.exponentBox.Text)
    If Me.showVarBox.Value Then
        Worksheets(glob.name).Visible = xlSheetVisible
    Else
        Worksheets(glob.name).Visible = xlSheetHidden
    End If
End Sub
Private Sub showVals1()
    On Error Resume Next
    Me.recnumBox.Value = 1
    Me.SpinButton.min = 1
    Me.tieNe2Be.Visible = (TabStrip.SelectedItem.Caption = "21Ne")
    Me.Label51.Visible = (TabStrip.SelectedItem.Caption = "21Ne")
    Me.P21Ne10Be.Visible = (TabStrip.SelectedItem.Caption = "21Ne")
    Select Case TabStrip.SelectedItem.Caption
        Case Is = "10Be"
            Me.SpinButton.Max = glob.n10BeCals
            Me.avgPBox.Text = Format(glob.P10Be, "#.00")
        Case Is = "26Al"
            Me.SpinButton.Max = glob.n26AlCals
            Me.avgPBox.Text = Format(glob.P26Al, "#.00")
        Case Is = "21Ne"
            Me.SpinButton.Max = glob.n21NeCals
            Me.avgPBox.Text = Format(glob.P21Ne, "#.00")
            Me.P21Ne10Be.Text = Format(glob.P21Ne10Be, "#.00")
        Case Is = "3He"
            Me.SpinButton.Max = glob.n3HeCals
            Me.avgPBox.Text = Format(glob.P3He, "#.00")
        Case Is = "36Cl"
            Me.SpinButton.Max = glob.n36ClCals
            Me.avgPBox.Text = Format(glob.P36Cl, "#.00")
        Case Is = "14C"
            Me.SpinButton.Max = glob.n14CCals
            Me.avgPBox.Text = Format(glob.P14C, "#.00")
    End Select
    Me.numrecBox.Value = Me.SpinButton.Max
    Me.ComboBox1.Value = glob.Scaling
    addRecordButton.Caption = "Add"
    reminder.Visible = False
    Call showRec(TabStrip.SelectedItem.Caption, 1)
End Sub
Private Sub showVals2()
    On Error Resume Next
    Me.L0Box.Text = CStr(glob.L0)
    Me.L1Box.Text = CStr(glob.L1)
    Me.L2Box.Text = CStr(glob.L2)
    Me.L3Box.Text = CStr(glob.L3)
    Me.F26Al0Box.Text = CStr(glob.F26Al0)
    Me.F26Al1Box.Text = CStr(glob.F26Al1)
    Me.F26Al2Box.Text = CStr(glob.F26Al2)
    Me.F26Al3Box.Text = CStr(glob.F26Al3)
    Me.F10Be0Box.Text = CStr(glob.F10Be0)
    Me.F10Be1Box.Text = CStr(glob.F10Be1)
    Me.F10Be2Box.Text = CStr(glob.F10Be2)
    Me.F10Be3Box.Text = CStr(glob.F10Be3)
    Me.F21Ne0Box.Text = CStr(glob.F21Ne0)
    Me.F21Ne1Box.Text = CStr(glob.F21Ne1)
    Me.F21Ne2Box.Text = CStr(glob.F21Ne2)
    Me.F21Ne3Box.Text = CStr(glob.F21Ne3)
    Me.F36Cl0Box.Text = CStr(glob.F36Cl0)
    Me.F36Cl1Box.Text = CStr(glob.F36Cl1)
    Me.F36Cl2Box.Text = CStr(glob.F36Cl2)
    Me.F36Cl3Box.Text = CStr(glob.F36Cl3)
    Me.F3He0Box.Text = CStr(glob.F3He0)
    Me.F3He1Box.Text = CStr(glob.F3He1)
    Me.F3He2Box.Text = CStr(glob.F3He2)
    Me.F3He3Box.Text = CStr(glob.F3He3)
    Me.F14C0Box.Text = CStr(glob.F14C0)
    Me.F14C1Box.Text = CStr(glob.F14C1)
    Me.F14C2Box.Text = CStr(glob.F14C2)
    Me.F14C3Box.Text = CStr(glob.F14C3)
    Me.RhoBox.Text = CStr(glob.rho)
    Me.L26AlBox.Text = CStr(glob.L26Al)
    Me.L10BeBox.Text = CStr(glob.L10Be)
    Me.L36ClBox.Text = CStr(glob.L36Cl)
    Me.L14CBox.Text = CStr(glob.L14C)
    If glob.PC Then
        Me.L26AlBox.Text = Format(glob.L26Al, "###E+0")
        Me.L10BeBox.Text = Format(glob.L10Be, "###E+0")
        Me.L36ClBox.Text = Format(glob.L36Cl, "###E+0")
        Me.L14CBox.Text = Format(glob.L14C, "###E+0")
    End If
End Sub
Private Sub showVals3()
    On Error Resume Next
    Me.ToBox.Text = CStr(glob.T_o)
    Me.BoBox.Text = CStr(glob.B_o)
    Me.PoBox.Text = CStr(glob.P_o)
    Me.MM0Box.Text = CStr(glob.MM0)
    Me.exponentBox.Text = CStr(glob.exponent)
End Sub
Private Sub HelpButton1_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Please enter the following information:" & vbCrLf & vbCrLf
    Msg = Msg & "N: TCN concentration in atoms per gram for: " & vbCrLf
    Msg = Msg & "        - SiO2 (for 10Be, 26Al, 21Ne and 14C)" & vbCrLf
    Msg = Msg & "        - (Fe,Mg)SiO4 (for 3He)" & vbCrLf
    Msg = Msg & "        - Ca (for 36Cl) " & vbCrLf & vbCrLf
    Msg = Msg & "Latitude: depending on the scaling model, this equals:" & vbCrLf
    Msg = Msg & "        - geomagnetic latitude in degrees (Lal, Stone)" & vbCrLf
    Msg = Msg & "        - geomagnetic inclination in degrees (Dunai)" & vbCrLf
    Msg = Msg & "        - cutoff rigidity in GV (Desilets et al)" & vbCrLf & vbCrLf
    Msg = Msg & "Elevation: depending on the scaling model, this equals:" & vbCrLf
    Msg = Msg & "        - altitude in m (Lal)" & vbCrLf
    Msg = Msg & "        - atmospheric pressure in mbar (Stone)" & vbCrLf
    Msg = Msg & "        - atmospheric depth in g/cm3 (Dunai, Desilets et al)" & vbCrLf & vbCrLf
    Msg = Msg & "Age: Independent age estimate for the calibration site" & vbCrLf & vbCrLf
    Msg = Msg & "P: TCN production rate at SLHL, calculated by CosmoCalc" & vbCrLf
    Msg = Msg & "          from N, Latitude, Elevation and Age." & vbCrLf & vbCrLf
    Msg = Msg & "CosmoCalc calculates the SLHL production rates implicitly" & vbCrLf
    Msg = Msg & "by default, but it is also possible to define them" & vbCrLf
    Msg = Msg & "explicitly in the 'Average P at SLHL' box and clicking 'Enter'" & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub HelpButton2_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "CosmoCalc uses the approach of Granger and Smith (2000) to" & vbCrLf
    Msg = Msg & "approximate the production rate function with a set of " & vbCrLf
    Msg = Msg & "four exponentials, governed by the following parameters:" & vbCrLf & vbCrLf
    Msg = Msg & "F0(n) =  relative SLHL surface production by neutrons" & vbCrLf
    Msg = Msg & "F1(m) =  relative SLHL surface production by slow muons (1st exponential)" & vbCrLf
    Msg = Msg & "F2(m) =  relative SLHL surface production by slow muons (2nd exponential)" & vbCrLf
    Msg = Msg & "F3(m) =  relative SLHL surface production by fast muons" & vbCrLf & vbCrLf
    Msg = Msg & "L0(n) =  neutron attenuation length" & vbCrLf
    Msg = Msg & "L1(m) =  slow muon attenuation length (1st exponential)" & vbCrLf
    Msg = Msg & "L2(m) =  slow muon attenuation length (2nd exponential)" & vbCrLf
    Msg = Msg & "L3(m) =  fast muon attenuation length" & vbCrLf & vbCrLf
    Msg = Msg & "Note: no (epi)thermal neutrons are accounted for." & vbCrLf
    Msg = Msg & "36Cl calculations may therefore be inaccurate" & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub HelpButton3_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "- Sea level temperature, thermal lapse rate and the" & vbCrLf
    Msg = Msg & "    atmospheric pressure at sea level are used to" & vbCrLf
    Msg = Msg & "    calculate the atmospheric depths required for" & vbCrLf
    Msg = Msg & "    Dunai and Desilets scaling" & vbCrLf & vbCrLf
    Msg = Msg & "- Magnetic field intensity is needed for the" & vbCrLf
    Msg = Msg & "    calculation of cutoff rigidities" & vbCrLf & vbCrLf
    Msg = Msg & "- Recommended values of the exponent used for" & vbCrLf
    Msg = Msg & "    topographic shielding corrections are 2.3 and" & vbCrLf
    Msg = Msg & "    3.5, respectively (Staudacher et al., 1993)" & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
