VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_numcols As Integer
Private m_numrows As Long
Private m_mycols() As String
Private m_rngname As String
Private m_shtname As String
Public Sub SetProperties(ByVal rng As String, Optional sheetName As Variant)
    On Error Resume Next:
    m_rngname = Application.WorksheetFunction.Substitute(rng, ";", ",")
    If IsMissing(sheetName) Then
        m_shtname = ActiveSheet.name
    Else
        m_shtname = sheetName
    End If
    m_numcols = 0
    numareas = Range(m_rngname).Areas.Count
    ' count the number of columns
    For i = 1 To numareas
        For j = 1 To Range(m_rngname).Areas(i).Columns.Count
            m_numcols = m_numcols + 1
        Next j
    Next i
    ReDim m_mycols(1 To m_numcols)
    k = 0
    ' store the ranges of all columns
    For i = 1 To numareas
        For j = 1 To Range(m_rngname).Areas(i).Columns.Count
            k = k + 1
            m_mycols(k) = Range(m_rngname).Areas(i).Columns(j).Address
        Next j
    Next i
    m_numrows = getNumRows(m_rngname)
End Sub
Property Get sheet() As String
    sheet = m_shtname
End Property
Property Get name() As String
    name = m_rngname
End Property
Property Get numcols() As Integer
    numcols = m_numcols
End Property
Property Get numRows() As Long
    numRows = m_numrows
End Property
Property Get nonEmptyCols(ByVal rownum As Integer) As Integer
    nonEmptyCols = 0
    For colnum = 1 To m_numcols
        If Worksheets(m_shtname).Range(m_mycols(colnum)).Cells(rownum).Value = "" Then
            Exit For
        Else
            nonEmptyCols = nonEmptyCols + 1
        End If
    Next colnum
End Property
Property Get ColumnAddress(ByVal colnum As Integer) As String
    ColumnAddress = Worksheets(m_shtname).m_mycols(colnum).Address
End Property
Property Get CellAddress(ByVal rownum As Integer, ByVal colnum As Integer) As String
    CellAddress = Worksheets(m_shtname).Range(m_mycols(colnum)).Cells(rownum).Address
End Property
Property Get CellValue(ByVal rownum As Integer, ByVal colnum As Integer) As Variant
    CellValue = Worksheets(m_shtname).Range(m_mycols(colnum)).Cells(rownum).Value
End Property
Public Sub SetCellValue(ByVal rownum As Integer, ByVal colnum As Integer, ByVal val As Variant)
    Worksheets(m_shtname).Range(m_mycols(colnum)).Cells(rownum).Value = val
End Sub
Private Function getNumRows(rng As String) As Long
    Dim LastRow&
    On Error Resume Next
    LastRow& = Worksheets(m_shtname).Range(rng).Cells.Find(What:="*", _
        SearchDirection:=xlPrevious, SearchOrder:=xlByRows).row
    ' if the user selected a rectangular range
    If LastRow& > Worksheets(m_shtname).Range(rng).Rows.Count Then
        getNumRows = Worksheets(m_shtname).Range(m_rngname).Rows.Count
    Else ' if they selected entire columns
        getNumRows = LastRow&
    End If
End Function
