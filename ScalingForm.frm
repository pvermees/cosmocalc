VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScalingForm 
   Caption         =   "Scaling Factors"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   -105
   ClientWidth     =   5055
   OleObjectBlob   =   "ScalingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScalingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox_Change()
    Select Case Me.ComboBox
        Case Is = "Lal"
            Me.Label.Caption = "lat (deg) - elev (m)"
        Case Is = "Stone"
            Me.Label.Caption = "lat (deg) - pressure (mbar)"
        Case Is = "Dunai"
            Me.Label.Caption = "incl (deg) - depth (g/cm3)"
        Case Is = "Desilets & Zreda (2003)"
            Me.Label.Caption = "Rc (GV) - depth (g/cm3)"
        Case Is = "Desilets et al (2006)"
            Me.Label.Caption = "Rc (GV) - depth (g/cm3)"
    End Select
End Sub

Private Sub userform_initialize()
    On Error Resume Next
    With Me.ComboBox
        .AddItem "Lal"
        .AddItem "Stone"
        .AddItem "Dunai"
        .AddItem "Desilets & Zreda (2003)"
        .AddItem "Desilets et al (2006)"
        .Value = glob.Scaling
    End With
    With Me.ComboBox1
        .AddItem "26Al"
        .AddItem "10Be"
        .AddItem "21Ne"
        .AddItem "3He"
        .AddItem "36Cl"
        .AddItem "14C"
        .Value = "10Be"
    End With
    RefEdit.Value = Selection.Address
End Sub
Private Sub OKButton_Click()
    On Error Resume Next
    Dim theRange As MyRange
    Dim oldScaling As String, newScaling As String
    oldScaling = glob.Scaling
    newScaling = Me.ComboBox.Value
    If oldScaling <> newScaling Then
        glob.Scaling = newScaling
        ' change P_SLHL according to the new scaling factors
        Call glob.ConvertAllP(oldScaling, newScaling)
    End If
    Set theRange = New MyRange
    Call theRange.SetProperties(Me.RefEdit.Value)
    Dim nucl As MyNuclide
    Set nucl = New MyNuclide
    Call nucl.SetProperties(ComboBox1.Value)
    If ComboBox.Value = "Lal" Then
        Call LalScaling(theRange)
    ElseIf ComboBox.Value = "Stone" Then
        Call StoneScaling(theRange, nucl)
    ElseIf ComboBox.Value = "Dunai" Then
        Call DunaiScaling(theRange, nucl)
    ElseIf ComboBox.Value = "Desilets & Zreda (2003)" Then
        Call DesiletScaling(2003, theRange, nucl)
    ElseIf ComboBox.Value = "Desilets et al (2006)" Then
        Call DesiletScaling(2006, theRange, nucl)
    End If
    Set theRange = Nothing
    Set nucl = Nothing
    Unload Me
End Sub
Private Sub HelpButton_Click()
    Dim Msg As String
    Msg = ""
    Msg = Msg & "Select a range of cells, two columns wide. " & vbCrLf
    Msg = Msg & "          x(1)  y(1)" & vbCrLf
    Msg = Msg & "           :     :  " & vbCrLf
    Msg = Msg & "          x(n)  y(n)" & vbCrLf
    Msg = Msg & "with x and y respectively:" & vbCrLf
    Msg = Msg & "  - latitude and elevation (Lal)" & vbCrLf
    Msg = Msg & "  - latitude and pressure (Stone)" & vbCrLf
    Msg = Msg & "  - magnetic inclination and atmospheric depth (Dunai)" & vbCrLf
    Msg = Msg & "  - cut-off rigidity and atmospheric depth (Desilets et al)" & vbCrLf & vbCrLf
    MsgBox Msg, vbInformation, APPNAME
End Sub
Private Sub CommandButton1_Click()
    Unload Me
End Sub
