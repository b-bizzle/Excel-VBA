Attribute VB_Name = "Colour_Cycle"
Option Explicit

Sub ColorCycle()
Attribute ColorCycle.VB_ProcData.VB_Invoke_Func = "q\n14"

    If Selection.Cells.Count > 1 Then
        With Selection.SpecialCells(xlCellTypeVisible).Interior
            If .colorindex = 2 Or .colorindex = -4142 Or .Color = 16777215 Or .colorindex = xlNone Then
                .colorindex = 6
            ElseIf .colorindex = 6 Then
                .colorindex = 3
            ElseIf .colorindex = 3 Then
                .colorindex = 46
            ElseIf .colorindex = 46 Then
                .colorindex = 35
            ElseIf .colorindex = 35 Then
                .colorindex = 36
            ElseIf .colorindex = 36 Then
                .colorindex = 37
            ElseIf .colorindex = 37 Then
                .colorindex = 15
            Else
                .colorindex = xlNone
            End If
        End With
    Else
        With Selection.Interior
            If .colorindex = 2 Or .colorindex = -4142 Or .Color = 16777215 Or .colorindex = xlNone Then
                .colorindex = 6
            ElseIf .colorindex = 6 Then
                .colorindex = 3
            ElseIf .colorindex = 3 Then
                .colorindex = 46
            ElseIf .colorindex = 46 Then
                .colorindex = 35
            ElseIf .colorindex = 35 Then
                .colorindex = 36
            ElseIf .colorindex = 36 Then
                .colorindex = 37
            ElseIf .colorindex = 37 Then
                .colorindex = 15
            Else
                .colorindex = xlNone
            End If
        End With
    End If

    
End Sub


Sub colorindex()

    Debug.Print Selection.Interior.colorindex

End Sub

Sub Reset_Colours()

    ActiveWorkbook.ResetColors
    
End Sub

