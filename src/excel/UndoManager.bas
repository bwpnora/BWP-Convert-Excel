Attribute VB_Name = "UndoManager"
Option Explicit

' ==============================================================================
' UndoManager.bas
' Transactional 2-phase 1-level Undo Manager for BWPConvertTTNVN.
' Features:
' 1. Guard threshold: MAX_UNDO_CELLS = 10000.
' 2. 2-phase commit: Staged buffer is only committed on successful batch completion.
' 3. Restoration: Restores .Formula for formulas, .Value2 for constants, ClearContents for blank cells.
' 4. Safe UI invalidation for ribbon button state transitions.
' Pure 7-bit ASCII safe source code.
' ==============================================================================

Public Const MAX_UNDO_CELLS As Long = 10000

Private Type VnUndoSnapshot
    WbName As String
    WsName As String
    Address As String
    RowCount As Long
    ColCount As Long
    Values As Variant
    Formulas As Variant
    HasFormula() As Boolean
    IsEmptyCell() As Boolean
    IsActive As Boolean
End Type

Private mStaged As VnUndoSnapshot
Private mCommitted As VnUndoSnapshot

' ==============================================================================
' Phase 1: Prepare Undo Snapshot (Staged Buffer)
' ==============================================================================

Public Sub PrepareUndoSnapshot(ByVal TargetRange As Range)
    On Error GoTo ErrHandler
    
    ' Always reset staged buffer first
    DiscardStagedSnapshot
    
    If TargetRange Is Nothing Then Exit Sub
    If TargetRange.Areas.Count <> 1 Then Exit Sub
    
    Dim totalCells As Double
    totalCells = CDbl(TargetRange.Rows.Count) * CDbl(TargetRange.Columns.Count)
    If totalCells > CDbl(MAX_UNDO_CELLS) Then
        ClearUndo
        Exit Sub
    End If
    
    Dim rCount As Long, cCount As Long
    rCount = TargetRange.Rows.Count
    cCount = TargetRange.Columns.Count
    
    mStaged.WbName = TargetRange.Parent.Parent.Name
    mStaged.WsName = TargetRange.Parent.Name
    mStaged.Address = TargetRange.Address
    mStaged.RowCount = rCount
    mStaged.ColCount = cCount
    
    ReDim mStaged.HasFormula(1 To rCount, 1 To cCount)
    ReDim mStaged.IsEmptyCell(1 To rCount, 1 To cCount)
    
    If rCount = 1 And cCount = 1 Then
        Dim valArr(1 To 1, 1 To 1) As Variant
        Dim fmlaArr(1 To 1, 1 To 1) As Variant
        
        valArr(1, 1) = TargetRange.Value2
        fmlaArr(1, 1) = TargetRange.Formula
        
        mStaged.Values = valArr
        mStaged.Formulas = fmlaArr
        
        mStaged.HasFormula(1, 1) = CBool(TargetRange.HasFormula)
        mStaged.IsEmptyCell(1, 1) = IsEmpty(valArr(1, 1)) And Not mStaged.HasFormula(1, 1)
    Else
        mStaged.Values = TargetRange.Value2
        mStaged.Formulas = TargetRange.Formula
        
        Dim r As Long, c As Long
        Dim rngHasFmla As Variant
        rngHasFmla = TargetRange.HasFormula
        
        If Not IsNull(rngHasFmla) Then
            If CBool(rngHasFmla) Then
                For r = 1 To rCount
                    For c = 1 To cCount
                        mStaged.HasFormula(r, c) = True
                        mStaged.IsEmptyCell(r, c) = False
                    Next c
                Next r
            Else
                For r = 1 To rCount
                    For c = 1 To cCount
                        mStaged.HasFormula(r, c) = False
                        mStaged.IsEmptyCell(r, c) = IsEmpty(mStaged.Values(r, c))
                    Next c
                Next r
            End If
        Else
            ' Mixed formulas and constants
            For r = 1 To rCount
                For c = 1 To cCount
                    mStaged.HasFormula(r, c) = False
                    mStaged.IsEmptyCell(r, c) = IsEmpty(mStaged.Values(r, c))
                Next c
            Next r
            
            Dim fmlaCells As Range
            Set fmlaCells = Nothing
            On Error Resume Next
            Set fmlaCells = TargetRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo ErrHandler
            
            If Not (fmlaCells Is Nothing) Then
                Dim cell As Range
                Dim baseRow As Long, baseCol As Long
                baseRow = TargetRange.Row
                baseCol = TargetRange.Column
                For Each cell In fmlaCells.Cells
                    Dim cellR As Long, cellC As Long
                    cellR = cell.Row - baseRow + 1
                    cellC = cell.Column - baseCol + 1
                    If cellR >= 1 And cellR <= rCount And cellC >= 1 And cellC <= cCount Then
                        mStaged.HasFormula(cellR, cellC) = True
                        mStaged.IsEmptyCell(cellR, cellC) = False
                    End If
                Next cell
            End If
        End If
    End If
    
    mStaged.IsActive = True
    Exit Sub

ErrHandler:
    DiscardStagedSnapshot
End Sub

' ==============================================================================
' Phase 2: Commit or Discard Staged Buffer
' ==============================================================================

Public Sub CommitUndoSnapshot()
    If Not mStaged.IsActive Then Exit Sub
    
    mCommitted.WbName = mStaged.WbName
    mCommitted.WsName = mStaged.WsName
    mCommitted.Address = mStaged.Address
    mCommitted.RowCount = mStaged.RowCount
    mCommitted.ColCount = mStaged.ColCount
    mCommitted.Values = mStaged.Values
    mCommitted.Formulas = mStaged.Formulas
    mCommitted.HasFormula = mStaged.HasFormula
    mCommitted.IsEmptyCell = mStaged.IsEmptyCell
    mCommitted.IsActive = True
    
    mStaged.IsActive = False
    
    InvalidateUndoUI
End Sub

Public Sub DiscardStagedSnapshot()
    mStaged.IsActive = False
    mStaged.Values = Empty
    mStaged.Formulas = Empty
    Erase mStaged.HasFormula
    Erase mStaged.IsEmptyCell
    ' Note: mCommitted is intentionally preserved intact.
End Sub

' ==============================================================================
' Undo Execution & Query
' ==============================================================================

Public Function ExecuteUndo() As Boolean
    On Error GoTo ErrHandler
    ExecuteUndo = False
    
    If Not mCommitted.IsActive Then Exit Function
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Nothing
    On Error Resume Next
    Set wb = Application.Workbooks(mCommitted.WbName)
    On Error GoTo ErrHandler
    If wb Is Nothing Then Exit Function
    
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(mCommitted.WsName)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Function
    
    Dim targetRng As Range
    Set targetRng = ws.Range(mCommitted.Address)
    If targetRng Is Nothing Then Exit Function
    
    ' State preservation
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevEnableEvents As Boolean
    
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim r As Long, c As Long
    Dim rCount As Long, cCount As Long
    rCount = mCommitted.RowCount
    cCount = mCommitted.ColCount
    
    For r = 1 To rCount
        For c = 1 To cCount
            Dim cell As Range
            If rCount = 1 And cCount = 1 And targetRng.MergeCells Then
                Set cell = targetRng.MergeArea.Cells(1, 1)
            Else
                Set cell = targetRng.Cells(r, c)
            End If
            
            If mCommitted.IsEmptyCell(r, c) Then
                cell.ClearContents
            ElseIf mCommitted.HasFormula(r, c) Then
                cell.Formula = mCommitted.Formulas(r, c)
            Else
                cell.Value2 = mCommitted.Values(r, c)
            End If
        Next c
    Next r
    
    ' Restore state
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    
    ' Clear committed snapshot after restoration (1-level undo)
    ClearUndo
    
    ExecuteUndo = True
    Exit Function

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    On Error GoTo 0
    ExecuteUndo = False
End Function

Public Sub ClearUndo()
    mStaged.IsActive = False
    mStaged.Values = Empty
    mStaged.Formulas = Empty
    Erase mStaged.HasFormula
    Erase mStaged.IsEmptyCell
    
    mCommitted.IsActive = False
    mCommitted.Values = Empty
    mCommitted.Formulas = Empty
    Erase mCommitted.HasFormula
    Erase mCommitted.IsEmptyCell
    
    InvalidateUndoUI
End Sub

Public Function IsUndoAvailable() As Boolean
    IsUndoAvailable = mCommitted.IsActive
End Function

' ==============================================================================
' UI Notification
' ==============================================================================

Private Sub InvalidateUndoUI()
    On Error Resume Next
    Application.Run "InvalidateUndoButton"
    On Error GoTo 0
End Sub
