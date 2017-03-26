VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConverterForm 
   Caption         =   "Converters"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   -540
   ClientWidth     =   6120
   OleObjectBlob   =   "ConverterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConverterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub userform_initialize()
    With ComboBoxElevFrom
        .AddItem "Elevation (m)"
        .AddItem "Pressure (mbar) standard atmosphere"
        .AddItem "Pressure (mbar) Antarctica"
        .AddItem "Atmospheric depth (g/cm2)"
    End With
    With ComboBoxElevTo
        .AddItem "Elevation (m)"
        .AddItem "Pressure (mbar) standard atmosphere"
        .AddItem "Pressure (mbar) Antarctica"
        .AddItem "Atmospheric depth (g/cm2)"
    End With
    With ComboBoxLatFrom
        .AddItem "x.x deg"
        .AddItem "x deg y.y min"
        .AddItem "x deg y min z sec"
        .AddItem "Inclination (deg)"
        .AddItem "Cutoff rigidity (GV)"
    End With
    With ComboBoxLatTo
        .AddItem "x.x deg"
        .AddItem "x deg y.y min"
        .AddItem "x deg y min z sec"
        .AddItem "Inclination (deg)"
        .AddItem "Cutoff rigidity (GV)"
    End With
    ElevationRefEdit.Value = Selection.Address
    LatitudeRefEdit.Value = Selection.Address
End Sub
Private Sub CancelButton_click()
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
Private Sub ConvertElevButton_Click()
    On Error Resume Next
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.ElevationRefEdit.Value)
    Select Case ComboBoxElevFrom.Value
    Case Is = "Elevation (m)"
        Select Case ComboBoxElevTo.Value
        Case Is = "Elevation (m)"
            ' nothing to do
        Case Is = "Pressure (mbar) standard atmosphere"
            Call zToP(theRange)
        Case Is = "Pressure (mbar) Antarctica"
            Call zToAntP(theRange)
        Case Is = "Atmospheric depth (g/cm2)"
            Call zToD(theRange)
        End Select
    Case Is = "Pressure (mbar) standard atmosphere"
        Select Case ComboBoxElevTo.Value
        Case Is = "Elevation (m)"
            Call pToZ(theRange)
        Case Is = "Pressure (mbar) standard atmosphere"
            ' nothing to do
        Case Is = "Pressure (mbar) Antarctica"
            Call pToAntP(theRange)
        Case Is = "Atmospheric depth (g/cm2)"
            Call pToD(theRange)
        End Select
    Case Is = "Pressure (mbar) Antarctica"
        Select Case ComboBoxElevTo.Value
        Case Is = "Elevation (m)"
            Call AntPtoZ(theRange)
        Case Is = "Pressure (mbar) standard atmosphere"
            Call AntPtoP(theRange)
        Case Is = "Pressure (mbar) Antarctica"
            ' nothing to do
        Case Is = "Atmospheric depth (g/cm2)"
            Call AntPtoD(theRange)
        End Select
    Case Is = "Atmospheric depth (g/cm2)"
        Select Case ComboBoxElevTo.Value
        Case Is = "Elevation (m)"
            Call dToZ(theRange)
        Case Is = "Pressure (mbar) standard atmosphere"
            Call dToP(theRange)
        Case Is = "Pressure (mbar) Antarctica"
            Call dToAntP(theRange)
        Case Is = "Atmospheric depth (g/cm2)"
            ' nothing to do
        End Select
    End Select
    Set theRange = Nothing
    Unload Me
End Sub
Private Sub ConvertLatLonButton_Click()
    On Error Resume Next
    Dim theRange As MyRange
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.LatitudeRefEdit.Value)
    Select Case ComboBoxLatFrom.Value
    Case Is = "x.x deg"
        Select Case ComboBoxLatTo.Value
        Case Is = "x.x deg"
            ' nothing to do
        Case Is = "x deg y.y min"
            Call degToDegMin(theRange)
        Case Is = "x deg y min z sec"
            Call degToDegMinSec(theRange)
        Case Is = "Inclination (deg)"
            Call degToI(theRange)
        Case Is = "Cutoff rigidity (GV)"
            Call degToRc(theRange)
        End Select
    Case Is = "x deg y.y min"
        Select Case ComboBoxLatTo.Value
        Case Is = "x.x deg"
            Call degMinToDeg(theRange)
        Case Is = "x deg y.y min"
            ' nothing to do
        Case Is = "x deg y min z sec"
            Call degMinToDegMinSec(theRange)
        Case Is = "Inclination (deg)"
            If glob.Replace Then
                Call degMinToDeg(theRange)
                Call degToI(theRange)
            Else
                Call degMinToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 2)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 2)).Offset(0, 1).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToI(theRange)
                glob.Replace = False
            End If
        Case Is = "Cutoff rigidity (GV)"
            If glob.Replace Then
                Call degMinToDeg(theRange)
                Call degToRc(theRange)
            Else
                Call degMinToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 2)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 2)).Offset(0, 1).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToRc(theRange)
                glob.Replace = False
            End If
        End Select
    Case Is = "x deg y min z sec"
        Select Case ComboBoxLatTo.Value
        Case Is = "x.x deg"
            Call degMinSecToDeg(theRange)
        Case Is = "x deg y.y min"
            Call degMinSecToDegMin(theRange)
        Case Is = "x deg y min z sec"
            ' nothing to do
        Case Is = "Inclination (deg)"
            If glob.Replace Then
                Call degMinSecToDeg(theRange)
                Call degToI(theRange)
            Else
                Call degMinSecToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 3)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 3)).Offset(0, 1).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToI(theRange)
                glob.Replace = False
            End If
        Case Is = "Cutoff rigidity (GV)"
            If glob.Replace Then
                Call degMinSecToDeg(theRange)
                Call degToRc(theRange)
            Else
                Call degMinSecToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 3)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 3)).Offset(0, 1).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToRc(theRange)
                glob.Replace = False
            End If
        End Select
    Case Is = "Inclination (deg)"
        Select Case ComboBoxLatTo.Value
        Case Is = "x.x deg"
            Call iToDeg(theRange)
        Case Is = "x deg y.y min"
            If glob.Replace Then
                Call iToDeg(theRange)
                Call degToDegMin(theRange)
            Else
                Call iToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 1)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 1)).Offset(0, 2).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToDegMin(theRange)
                glob.Replace = False
            End If
        Case Is = "x deg y min z sec"
            If glob.Replace Then
                Call iToDeg(theRange)
                Call degToDegMinSec(theRange)
            Else
                Call iToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 1)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 1)).Offset(0, 3).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToDegMinSec(theRange)
                glob.Replace = False
            End If
        Case Is = "Inclination (deg)"
            ' nothing to do
        Case Is = "Cutoff rigidity (GV)"
            If glob.Replace Then
                Call iToDeg(theRange)
                Call degToRc(theRange)
            Else
                Call iToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 1)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 1)).Offset(0, 1).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToRc(theRange)
                glob.Replace = False
            End If
        End Select
    Case Is = "Cutoff rigidity (GV)"
        Select Case ComboBoxLatTo.Value
        Case Is = "x.x deg"
            Call RcToDeg(theRange)
        Case Is = "x deg y.y min"
            If glob.Replace Then
                Call RcToDeg(theRange)
                Call degToDegMin(theRange)
            Else
                Call RcToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 1)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 1)).Offset(0, 2).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToDegMin(theRange)
                glob.Replace = False
            End If
        Case Is = "x deg y min z sec"
            If glob.Replace Then
                Call RcToDeg(theRange)
                Call degToDegMinSec(theRange)
            Else
                Call RcToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 1)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 1)).Offset(0, 3).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToDegMinSec(theRange)
                glob.Replace = False
            End If
        Case Is = "Inclination (deg)"
            If glob.Replace Then
                Call RcToDeg(theRange)
                Call degToI(theRange)
            Else
                Call RcToDeg(theRange)
                firstCell = Range(theRange.CellAddress(1, 1)).Offset(0, 1).Address
                LastCell = Range(theRange.CellAddress(theRange.numRows, 1)).Offset(0, 1).Address
                Call theRange.SetProperties(Range(firstCell, LastCell).Address)
                glob.Replace = True
                Call degToI(theRange)
                glob.Replace = False
            End If
        Case Is = "Cutoff rigidity (GV)"
            ' nothing to do
        End Select
    End Select
    Set theRange = Nothing
    Unload Me
End Sub
Private Sub latOptionButton_Click()
    Call ConvertersOptionForm.Show
End Sub
Private Sub OptionButton_Click()
    Call ConvertersOptionForm.Show
End Sub
Private Sub HelpElevButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Select a range of cells, one column wide        " & vbCrLf & vbCrLf
    Msg = Msg & "          x(1)          " & vbCrLf
    Msg = Msg & "            :           " & vbCrLf
    Msg = Msg & "          x(n)          " & vbCrLf & vbCrLf
    Msg = Msg & "with the x-values elevation, pressure or" & vbCrLf
    Msg = Msg & "atmospheric depth values, depending on        " & vbCrLf
    Msg = Msg & "the radio-button selected in the 'To:'-field." & vbCrLf
    Msg = Msg & "The x-values are replaced by and converted" & vbCrLf
    Msg = Msg & "to the units selected in the 'From:'-field." & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub HelpLatLonButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Select a range of cells, one to three columns wide," & vbCrLf
    Msg = Msg & "depending on the radio-button selected in the" & vbCrLf
    Msg = Msg & "'From:'-field. These values are replaced by and" & vbCrLf
    Msg = Msg & "converted to the format selected in the 'To:'-field." & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
