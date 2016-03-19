Attribute VB_Name = "Module_Main"
Option Explicit
Option Private Module
Public Const APPNAME As String = "CosmoCalc"
Public Const pi As Double = 3.14159265358979
Public glob As MyGlobals
Sub Auto_Open()
'   Creates a new menu and adds menu items
    Dim Cap(1 To 7)
    Dim Mac(1 To 7)
    Dim MenuName As String
    Dim sht As Worksheet
    
    MenuName = "&CosmoCalc"
    
    Cap(1) = "Scaling"
    Mac(1) = "Scaling"
    Cap(2) = "Shielding"
    Mac(2) = "Shielding"
    Cap(3) = "Banana"
    Mac(3) = "Banana"
    Cap(4) = "Age/Erosion"
    Mac(4) = "Age"
    Cap(5) = "Converters"
    Mac(5) = "Converters"
    Cap(6) = "Settings"
    Mac(6) = "Settings"
    Cap(7) = "About..."
    Mac(7) = "About"

    On Error Resume Next
'   Delete the menu if it already exists
    MenuBars(xlWorksheet).Menus(MenuName).Delete
    
'   Add the menu
    MenuBars(xlWorksheet).Menus.Add Caption:=MenuName
    
'   Add the menu items
    With MenuBars(xlWorksheet).Menus(MenuName).MenuItems
        .Add Caption:=Cap(1), OnAction:=Mac(1)
        .Add Caption:=Cap(2), OnAction:=Mac(2)
        .Add Caption:=Cap(3), OnAction:=Mac(3)
        .Add Caption:=Cap(4), OnAction:=Mac(4)
        .Add Caption:=Cap(5), OnAction:=Mac(5)
        .Add Caption:=Cap(6), OnAction:=Mac(6)
        .Add Caption:="-"
        .Add Caption:=Cap(7), OnAction:=Mac(7)
    End With
    
End Sub
Private Sub Scaling()
    On Error Resume Next
    Set glob = New MyGlobals
    ScalingForm.Show
    Set glob = Nothing
End Sub
Private Sub Shielding()
    On Error Resume Next
    Set glob = New MyGlobals
    ShieldingForm.Show
    Set glob = Nothing
End Sub
Private Sub Age()
    On Error Resume Next
    Set glob = New MyGlobals
    AgeForm.Show
    Set glob = Nothing
End Sub
Private Sub Banana()
    On Error Resume Next
    Set glob = New MyGlobals
    BananaForm.Show
    Set glob = Nothing
End Sub
Private Sub Settings()
    On Error Resume Next
    Set glob = New MyGlobals
    SettingsForm.Show
    Set glob = Nothing
End Sub
Private Sub Converters()
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
Sub Auto_Close()
    Dim MenuName As String
    MenuName = "&CosmoCalc"
'   Delete the menu before closing
    On Error Resume Next
    MenuBars(xlWorksheet).Menus(MenuName).Delete
End Sub
