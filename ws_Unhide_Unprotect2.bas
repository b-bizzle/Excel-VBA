Attribute VB_Name = "ws_Unhide_Unprotect2"
Option Explicit

Sub Unhide_Unprotect()

Dim strPWord(1 To 10) As String
Dim lPNo As Long
Dim ws As Worksheet

'Password for OneNote Workstreams and Project Planning Toolkit = 'Random'

'Microsoft Excel
    strPWord(1) = "Random!?"
    strPWord(2) = "Random!!"
'Commitment Record
    strPWord(3) = "Random"
'Password for a workbook I created but cannot remember which one
    strPWord(4) = "Random"
'SASU Central Diary
    strPWord(5) = "Random"
'Demand workbook
    strPWord(6) = "Random"
'Ofsted SIF Data
    strPWord(7) = "Random"
'Legal Tracker
    strPWord(8) = "Random"
'UWMT
    strPWord(9) = "Random"
'SC Tool
    strPWord(10) = "Random"


    On Error Resume Next

        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible

            lPNo = 1
            
            For lPNo = 1 To UBound(strPWord) Step 1
                If ws.ProtectContents = False Then
                    Exit For
                Else
                    ws.Unprotect strPWord(lPNo)
                End If
            Next lPNo

            ws.Unprotect
        Next ws

    On Error GoTo 0

End Sub

