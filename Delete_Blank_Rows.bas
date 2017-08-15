Attribute VB_Name = "Delete_Blank_Rows"
Option Explicit

Public Sub DeleteBlankRows()

Dim lLastRow As Long
Dim lLastColumn As Long
Dim wSht As Worksheet

    If ActiveWorkbook Is Nothing Then
        MsgBox "No open workbook to delete rows from. Please open a workbook " & _
            "and try again", vbOKOnly Or vbInformation, "No Open Workbook!"
        GoTo Terminate
    End If
    
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With

    On Error GoTo ErrHnd

    Set wSht = Application.ActiveSheet

    With wSht
    
        lLastRow = .Cells.Find(What:="*", After:=.Cells(1), LookAt:=xlPart, _
            LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
        
        lLastColumn = .Cells.Find(What:="*", After:=.Cells(1), LookAt:=xlPart, _
            LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:= _
            xlPrevious, MatchCase:=False).Column
        
        .Range(.Cells(1, lLastColumn + 1), .Cells(lLastRow, lLastColumn + 1)). _
            Formula = "=IF(COUNTA(" & .Range(.Cells(1, 1), .Cells(1, lLastColumn)). _
            Address(RowAbsolute:=False, ColumnAbsolute:=False) & ")=0,"""",Row())"
        
        .Range(.Cells(1, lLastColumn + 1), .Cells(lLastRow, lLastColumn + 1)). _
            Value = .Range(.Cells(1, lLastColumn + 1), .Cells(lLastRow, _
            lLastColumn + 1)).Value
        
        .Cells.Columns.Sort Key1:=.Cells.Columns(lLastColumn + 1), Order1:= _
            xlAscending, HEADER:=xlNo, OrderCustom:=1, MatchCase:=False, _
            Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
        
        .Range(.Cells(1, lLastColumn + 1), .Cells(lLastRow, lLastColumn + 1)).Delete
        
        .Range("A1").Select
    
    End With

    GoTo Terminate

ErrHnd:
    MsgBox "Error number " & Err.Number & " encountered in DeleteBlankRows" & _
        vbCrLf & Err.Description, vbOKOnly Or vbInformation, "Error"

Terminate:
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub


