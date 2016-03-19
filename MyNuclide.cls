VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyNuclide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_name As String
Private m_L As Double
Private m_P As Double
Private m_F0 As Double
Private m_F1 As Double
Private m_F2 As Double
Private m_F3 As Double
Private m_S0 As Double
Private m_S1 As Double
Private m_S2 As Double
Private m_S3 As Double
Property Let name(ByVal newvalue As String)
    m_name = newvalue
End Property
Property Get name() As String
    name = m_name
End Property
Property Let L(ByVal newvalue As Double)
    m_L = newvalue
End Property
Property Get L() As Double
    L = m_L
End Property
Property Let P(ByVal newvalue As Double)
    m_P = newvalue
End Property
Property Get P() As Double
    P = m_P
End Property
Property Let F0(ByVal newvalue As Double)
    m_F0 = newvalue
End Property
Property Get F0() As Double
    F0 = m_F0
End Property
Property Let F1(ByVal newvalue As Double)
    m_F1 = newvalue
End Property
Property Get F1() As Double
    F1 = m_F1
End Property
Property Let F2(ByVal newvalue As Double)
    m_F2 = newvalue
End Property
Property Get F2() As Double
    F2 = m_F2
End Property
Property Let F3(ByVal newvalue As Double)
    m_F3 = newvalue
End Property
Property Get F3() As Double
    F3 = m_F3
End Property
Property Let S0(ByVal newvalue As Double)
    m_S0 = newvalue
End Property
Property Get S0() As Double
    S0 = m_S0
End Property
Property Let S1(ByVal newvalue As Double)
    m_S1 = newvalue
End Property
Property Get S1() As Double
    S1 = m_S1
End Property
Property Let S2(ByVal newvalue As Double)
    m_S2 = newvalue
End Property
Property Get S2() As Double
    S2 = m_S2
End Property
Property Let S3(ByVal newvalue As Double)
    m_S3 = newvalue
End Property
Property Get S3() As Double
    S3 = m_S3
End Property
Public Sub SetProperties(ByVal name As String)
    On Error Resume Next
    m_name = name
    m_S0 = 1
    m_S1 = 1
    m_S2 = 1
    m_S3 = 1
    Select Case m_name
        Case Is = "26Al"
            m_P = glob.P26Al
            m_L = glob.L26Al
            m_F0 = glob.F26Al0
            m_F1 = glob.F26Al1
            m_F2 = glob.F26Al2
            m_F3 = glob.F26Al3
        Case Is = "10Be"
            m_P = glob.P10Be
            m_L = glob.L10Be
            m_F0 = glob.F10Be0
            m_F1 = glob.F10Be1
            m_F2 = glob.F10Be2
            m_F3 = glob.F10Be3
        Case Is = "3He"
            m_P = glob.P3He
            m_L = glob.L3He
            m_F0 = glob.F3He0
            m_F1 = glob.F3He1
            m_F2 = glob.F3He2
            m_F3 = glob.F3He3
        Case Is = "21Ne"
            m_P = glob.P21Ne
            m_L = glob.L21Ne
            m_F0 = glob.F21Ne0
            m_F1 = glob.F21Ne1
            m_F2 = glob.F21Ne2
            m_F3 = glob.F21Ne3
        Case Is = "14C"
            m_P = glob.P14C
            m_L = glob.L14C
            m_F0 = glob.F14C0
            m_F1 = glob.F14C1
            m_F2 = glob.F14C2
            m_F3 = glob.F14C3
        Case Is = "36Cl"
            m_P = glob.P36Cl
            m_L = glob.L36Cl
            m_F0 = glob.F36Cl0
            m_F1 = glob.F36Cl1
            m_F2 = glob.F36Cl2
            m_F3 = glob.F36Cl3
    End Select
End Sub
Public Sub SetScaling(ByVal S As Double)
    x = getSx(S)
    m_S0 = Exp(-x / glob.L0)
    m_S1 = Exp(-x / glob.L1)
    m_S2 = Exp(-x / glob.L2)
    m_S3 = Exp(-x / glob.L3)
End Sub
Private Function getSx(ByVal S As Double) As Double
    PS = S
    x = 0 'initial guess (g/cm2)
    ' find root with Newton's method
    For i = 1 To 100
        PSest = m_F0 * Exp(-x / glob.L0) + m_F1 * Exp(-x / glob.L1) + _
                m_F2 * Exp(-x / glob.L2) + m_F3 * Exp(-x / glob.L3)
        If Abs((PSest - PS) / PS) < glob.Zero Then
            Exit For
        End If
        dPSdx = (m_F0 / glob.L0) * Exp(-x / glob.L0) + _
                (m_F1 / glob.L1) * Exp(-x / glob.L1) + _
                (m_F2 / glob.L2) * Exp(-x / glob.L2) + _
                (m_F3 / glob.L3) * Exp(-x / glob.L3)
        dx = (PSest - PS) / dPSdx
        x = x + dx
    Next i
    getSx = x
End Function
