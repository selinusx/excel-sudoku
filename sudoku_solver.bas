'==========================================================
' Sudoku for Excel – full module (Mac/Win compatible)
' Grid: A1:I9
' Provides: board setup, validation + conditional formatting,
'           example loader, clear, and a backtracking solver.
'==========================================================
Option Explicit

'---- config
Private Const START_ROW As Long = 1
Private Const START_COL As Long = 1
Private Const SIZE      As Long = 9

'=================== PUBLIC ENTRIES =======================
Public Sub SetupSudokuBoard()
    ' draws a 9x9 board at A1:I9, sizes cells, sets borders,
    ' and adds validation + conditional formatting rules
    Dim ws As Worksheet, r As Long, c As Long
    Set ws = ActiveSheet
    
    ' size cells
    For c = 1 To SIZE
        ws.Columns(START_COL + c - 1).ColumnWidth = 4
    Next c
    For r = 1 To SIZE
        ws.Rows(START_ROW + r - 1).RowHeight = 24
    Next r
    
    ' clear old content/format on grid
    With ws.Range(CellRef(1, 1), CellRef(SIZE, SIZE))
        .Clear
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline
    End With
    
    ' thick borders for 3x3 boxes
    Call DrawThickBoxBorders(ws)
    
    ' data validation 1..9 (allow blank)
    Call AddValidation(ws)
    
    ' conditional formatting: duplicates in row/col/3x3 box -> red
    Call AddConditionalFormatting(ws)
    
    MsgBox "Board ready at A1:I9. You can type numbers 1–9.", vbInformation
End Sub

Public Sub LoadExample()
    ' loads an easy puzzle ("" = blank)
    Dim arr As Variant
    arr = Array( _
        Array("", "", 3, "", 2, "", 6, "", ""), _
        Array(9, "", "", 3, "", 5, "", "", 1), _
        Array("", "", 1, 8, "", 6, 4, "", ""), _
        Array("", "", 8, 1, "", 2, 9, "", ""), _
        Array(7, "", "", "", "", "", "", "", 8), _
        Array("", "", 6, 7, "", 8, 2, "", ""), _
        Array("", "", 2, 6, "", 9, 5, "", ""), _
        Array(8, "", "", 2, "", 3, "", "", 9), _
        Array("", "", 5, "", 1, "", 3, "", "") _
    )
    
    Dim r As Long, c As Long
    For r = 1 To SIZE
        For c = 1 To SIZE
            CellRef(r, c).Value = arr(r - 1)(c - 1)
        Next c
    Next r
End Sub

Public Sub ClearGrid()
    Range(CellRef(1, 1), CellRef(SIZE, SIZE)).ClearContents
End Sub

Public Sub SolveSudoku()
    Dim ok As Boolean
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo SAFE_EXIT
    
    ok = SolveCell(1, 1)
    
SAFE_EXIT:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbExclamation
    ElseIf ok Then
        MsgBox "Solved", vbInformation
    Else
        MsgBox "No solution found (check puzzle).", vbExclamation
    End If
End Sub

'=================== CORE SOLVER ==========================
Private Function SolveCell(ByVal r As Long, ByVal c As Long) As Boolean
    ' move to next row when column exceeds 9
    If c > SIZE Then
        c = 1
        r = r + 1
    End If
    ' finished all rows => solved
    If r > SIZE Then
        SolveCell = True
        Exit Function
    End If
    
    If IsEmpty(CellRef(r, c).Value) Then
        Dim v As Long
        For v = 1 To 9
            If IsValid(r, c, v) Then
                CellRef(r, c).Value = v
                If SolveCell(r, c + 1) Then
                    SolveCell = True
                    Exit Function
                End If
                CellRef(r, c).ClearContents
            End If
        Next v
        SolveCell = False
    Else
        SolveCell = SolveCell(r, c + 1)
    End If
End Function

Private Function IsValid(ByVal r As Long, ByVal c As Long, ByVal v As Long) As Boolean
    Dim i As Long, j As Long
    
    ' row
    For j = 1 To SIZE
        If j <> c Then
            If CellRef(r, j).Value = v Then IsValid = False: Exit Function
        End If
    Next j
    ' col
    For i = 1 To SIZE
        If i <> r Then
            If CellRef(i, c).Value = v Then IsValid = False: Exit Function
        End If
    Next i
    ' 3x3 box
    Dim br As Long, bc As Long
    br = Int((r - 1) / 3) * 3 + 1
    bc = Int((c - 1) / 3) * 3 + 1
    For i = 0 To 2
        For j = 0 To 2
            If Not (br + i = r And bc + j = c) Then
                If CellRef(br + i, bc + j).Value = v Then
                    IsValid = False: Exit Function
                End If
            End If
        Next j
    Next i
    
    IsValid = True
End Function

'=================== BOARD HELPERS ========================
Private Sub DrawThickBoxBorders(ws As Worksheet)
    Dim r As Long, c As Long
    Dim rng As Range
    Set rng = ws.Range(CellRef(1, 1), CellRef(SIZE, SIZE))
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlHairline
    
    ' outer border
    With rng.Borders(xlEdgeLeft): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlEdgeTop): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    With rng.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    
    ' thick lines between boxes (after col 3 and 6; after row 3 and 6)
    For r = 1 To SIZE
        With CellRef(r, 3).Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With CellRef(r, 6).Borders(xlEdgeRight): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    Next r
    For c = 1 To SIZE
        With CellRef(3, c).Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
        With CellRef(6, c).Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Weight = xlMedium: End With
    Next c
End Sub

Private Sub AddValidation(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(CellRef(1, 1), CellRef(SIZE, SIZE))
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0
    rng.Validation.Add Type:=xlValidateWholeNumber, _
        AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=1, Formula2:=9
    With rng.Validation
        .IgnoreBlank = True
        .InputTitle = "Sudoku"
        .InputMessage = "1–9 gir (boş da olabilir)."
        .ErrorTitle = "Geçersiz giriş"
        .ErrorMessage = "Sadece 1–9 arası sayı gir."
    End With
End Sub

Private Sub AddConditionalFormatting(ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(CellRef(1, 1), CellRef(SIZE, SIZE))
    rng.FormatConditions.Delete
    
    ' column duplicate: =AND(A1<>"",COUNTIF(A$1:A$9,A1)>1)
    With rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(A1<>"""",COUNTIF(A$1:A$9,A1)>1)")
        .Interior.Color = RGB(255, 200, 200)
    End With
    ' row duplicate: =AND(A1<>"",COUNTIF($A1:$I1,A1)>1)
    With rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(A1<>"""",COUNTIF($A1:$I1,A1)>1)")
        .Interior.Color = RGB(255, 200, 200)
    End With
    ' 3x3 box duplicate using OFFSET around each cell
    ' =AND(A1<>"",COUNTIF(OFFSET(A1,-MOD(ROW(A1)-1,3),-MOD(COLUMN(A1)-1,3),3,3),A1)>1)
    With rng.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(A1<>"""",COUNTIF(OFFSET(A1,-MOD(ROW(A1)-1,3),-MOD(COLUMN(A1)-1,3),3,3),A1)>1)")
        .Interior.Color = RGB(255, 200, 200)
    End With
End Sub

Private Function CellRef(ByVal r As Long, ByVal c As Long) As Range
    Set CellRef = Cells(START_ROW + r - 1, START_COL + c - 1)
End Function
