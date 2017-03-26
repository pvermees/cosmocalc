Attribute VB_Name = "Module_Main"
Option Explicit
Option Private Module
Public Const APPNAME As String = "CosmoCalc"
Public Const pi As Double = 3.14159265358979
Public glob As MyGlobals
Public Sub ScalingModels()
    On Error Resume Next
    Set glob = New MyGlobals
    ScalingForm.Show
    Set glob = Nothing
End Sub
Public Sub Shielding()
    On Error Resume Next
    Set glob = New MyGlobals
    ShieldingForm.Show
    Set glob = Nothing
End Sub
Public Sub Age()
    On Error Resume Next
    Set glob = New MyGlobals
    AgeForm.Show
    Set glob = Nothing
End Sub
Public Sub Banana()
    On Error Resume Next
    Set glob = New MyGlobals
    BananaForm.Show
    Set glob = Nothing
End Sub
Public Sub Settings()
    On Error Resume Next
    Set glob = New MyGlobals
    SettingsForm.Show
    Set glob = Nothing
End Sub
Public Sub Converters()
    On Error Resume Next
    Set glob = New MyGlobals
    ConverterForm.Show
    Set glob = Nothing
End Sub
Public Sub DeleteRangeName(ByVal RangeName As String)
Dim rName As name
    For Each rName In ActiveWorkbook.Names
        If rName.name = RangeName Then
            rName.Delete
        End If
    Next rName
End Sub
Public Sub removeSheet(ByVal sheetName As String)
    Dim mysheet As Worksheet
    
    For Each mysheet In Worksheets
        If mysheet.name = sheetName Then
            Application.DisplayAlerts = False
            mysheet.Delete
            Application.DisplayAlerts = True
        End If
    Next mysheet
End Sub
Private Sub About()
    Set glob = New MyGlobals
    Dim Msg As String
    Msg = ""
    Msg = Msg & "The CosmoCalc add-in performs the following functions:" & vbCrLf & vbCrLf
    Msg = Msg & "- Calculate scaling factors (Lal, Stone, Dunai, Desilets)." & vbCrLf
    Msg = Msg & "- Calculate topographic, snow and self-shielding factors." & vbCrLf
    Msg = Msg & "- Generate banana-plots (currently only Al-Be and Ne-Be)." & vbCrLf
    Msg = Msg & "- Calculate exposure ages, erosion rates and burial ages." & vbCrLf
    Msg = Msg & "- Convert elevation to pressure, latitude to inclination, etc" & vbCrLf & vbCrLf
    Msg = Msg & "For user instructions, please check the help buttons of" & vbCrLf
    Msg = Msg & "the member functions. CosmoCalc " & glob.version & " was" & vbCrLf
    Msg = Msg & "created by Pieter Vermeesch for CRONUS-EU. The" & vbCrLf
    Msg = Msg & "latest version of the add-in can be downloaded from" & vbCrLf
    Msg = Msg & "the CosmoCalc website, along with a set of test data:" & vbCrLf & vbCrLf
    Msg = Msg & "http://cosmocalc.london-geochron.com" & vbCrLf & vbCrLf
    Msg = Msg & "Do you want to visit that site now?" & vbCrLf
    If MsgBox(Msg, vbInformation + vbYesNo, APPNAME) = vbYes Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink Address:="http://cosmocalc.london-geochron.com", NewWindow:=True
        On Error GoTo 0
    End If
    Set glob = Nothing
End Sub
