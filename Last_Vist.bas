Attribute VB_Name = "Last_Vist"
Option Explicit

Sub LastVisit()

Dim bFound As Boolean
Dim iWbkCount As Integer: iWbkCount = 0
Dim lINV06FirstRow As Long: lINV06FirstRow = 2
Dim lINV06LastRow As Long
Dim lINV06Row As Long
Dim lMax As Long
Dim lOffset As Long
Dim lRPT07FirstRow As Long: lRPT07FirstRow = 6
Dim lRPT07LastRow As Long
Dim lRPT07Row As Long
Dim rCell As Range
Dim wbkINV06 As Workbook
Dim wbkRPT07 As Workbook
Dim wbkTemp As Workbook
Dim wshtINV06 As Worksheet
Dim wshtRPT07 As Worksheet

    For Each wbkTemp In Workbooks
        If wbkTemp.Windows(1).Visible = True Then
            iWbkCount = iWbkCount + 1
            If wbkTemp.Name Like "*RPT07*" Then
                Set wbkRPT07 = wbkTemp
            ElseIf wbkTemp.Name Like "*INV06*" Then
                Set wbkINV06 = wbkTemp
            End If
        End If
    Next wbkTemp
    
    If iWbkCount <> 2 Then
        MsgBox "Please close all workbooks except for the INV06 and RPT07, then try again.", _
            vbOKOnly Or vbInformation, "Workbooks"
        Exit Sub
    End If
    
    If wbkRPT07 Is Nothing Or wbkINV06 Is Nothing Then
        MsgBox "Could not find the correct workbooks. Please open INV06 and RPT07, " & _
            "then try again.", vbOKOnly Or vbInformation, "Workbooks"
        Exit Sub
    End If
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    Set wshtINV06 = wbkINV06.Worksheets(1)
    Set wshtRPT07 = wbkRPT07.Worksheets(1)
    
    With wshtINV06
        lINV06LastRow = .Cells.Find(What:="*", After:=.Cells(1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
        .Range("A1:L" & lINV06LastRow - 1).Sort Key1:=.Range("A1"), Order1:=xlAscending, _
            Key2:=.Range("G1"), Order2:=xlDescending, HEADER:=xlYes
    End With
    
    With wshtRPT07
        lRPT07LastRow = .Cells.Find(What:="*", After:=.Cells(1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
    End With
            
    For lINV06Row = lINV06LastRow To lINV06FirstRow Step -1
        Application.StatusBar = "Process will be complete when counter reaches zero: " & _
            lINV06Row
        bFound = False
        If wshtINV06.Cells(lINV06Row, 6).Value = "Case Worker Visit" And _
            wshtINV06.Cells(lINV06Row, 2).Value <> "Unborn" Then
            GoTo SkipLoop
        End If
        For lRPT07Row = lRPT07FirstRow To lRPT07LastRow
            If wshtINV06.Cells(lINV06Row, 1).Value = wshtRPT07.Cells(lRPT07Row, 1).Value Then
                If wshtINV06.Cells(lINV06Row, 7).Value >= _
                    wshtRPT07.Cells(lRPT07Row, 17).Value Then
                    If wshtRPT07.Cells(lRPT07Row, 23).Value = "" Then
                        bFound = True
                        Exit For
                    Else
                        If wshtINV06.Cells(lINV06Row, 7).Value <= _
                            wshtRPT07.Cells(lRPT07Row, 23).Value Then
                            bFound = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next lRPT07Row
SkipLoop:
        If bFound = False Then
            wshtINV06.Cells(lINV06Row, 1).EntireRow.Delete
        Else
            If wshtINV06.Cells(lINV06Row, 1).Value = wshtINV06.Cells(lINV06Row + 1, 1).Value Then
                wshtINV06.Cells(lINV06Row + 1, 13).Value = _
                    wshtINV06.Cells(lINV06Row, 7).Value - wshtINV06.Cells(lINV06Row + 1, 7).Value
            
            End If
        End If
    Next lINV06Row
    
    With wshtINV06
        lINV06LastRow = .Cells.Find(What:="*", After:=.Cells(1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
        For Each rCell In .Range("M" & lINV06FirstRow & ":M" & lINV06LastRow)
            If rCell.Value = "" And rCell.Offset(1, 0).Value <> "" Then
                lOffset = 1
                lMax = 0
                Do Until rCell.Offset(lOffset, 0).Value = ""
                    If lMax < rCell.Offset(lOffset, 0).Value Then
                        lMax = rCell.Offset(lOffset, 0).Value
                    End If
                    lOffset = lOffset + 1
                Loop
                If lMax <= 28 Then
                    rCell.Value = "Y"
                Else
                    rCell.Value = "N"
                End If
            ElseIf rCell.Value = "" Then
                rCell.Value = "Y"
            End If
        Next rCell
    End With
    
    With wshtRPT07
        .Range(.Cells(lRPT07FirstRow, 20), .Cells(lRPT07LastRow, 20)).Formula = _
            "=VLOOKUP($A" & lRPT07FirstRow & ",'[" & wbkINV06.Name & "]" & wshtINV06.Name & _
            "'!$A$" & lINV06FirstRow & ":$M$" & lINV06LastRow & ",7,0)"
        .Range(.Cells(lRPT07FirstRow, 21), .Cells(lRPT07LastRow, 21)).Formula = _
            "=VLOOKUP($A" & lRPT07FirstRow & ",'[" & wbkINV06.Name & "]" & wshtINV06.Name & _
            "'!$A$" & lINV06FirstRow & ":$M$" & lINV06LastRow & ",13,0)"
        .Range(.Cells(lRPT07FirstRow, 22), .Cells(lRPT07LastRow, 22)).Formula = _
            "=VLOOKUP($A" & lRPT07FirstRow & ",'[" & wbkINV06.Name & "]" & wshtINV06.Name & _
            "'!$A$" & lINV06FirstRow & ":$M$" & lINV06LastRow & ",6,0)"
        .Range(.Cells(lRPT07FirstRow, 20), .Cells(lRPT07LastRow, 22)).Value = _
            .Range(.Cells(lRPT07FirstRow, 20), .Cells(lRPT07LastRow, 22)).Value
        .Range(.Cells(lRPT07FirstRow, 20), .Cells(lRPT07LastRow, 20)).NumberFormat = "dd/mm/yyyy"
    End With
            
    With Application
        .StatusBar = False
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With

End Sub

