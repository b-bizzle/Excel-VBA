Attribute VB_Name = "Protection2"
Option Explicit



Sub PasswordUnprotect()
'Password for OneNote Workstreams and Project Planning Toolkit = "Random"
'Password for UWMT VBA Project = "Random"
'Password for UWMT Archive = "Random"
'Password for Weekly Archive = "Random" - Not protected yet
'Password for Legal Tracer VBA = "Random"

Dim strPWord(1 To 10) As String
Dim lPNo As Long


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

    lPNo = 1
    
    For lPNo = 1 To UBound(strPWord) Step 1
        If ActiveSheet.ProtectContents = False Then
            Exit For
        Else
            ActiveSheet.Unprotect strPWord(lPNo)
        End If
    Next lPNo

    ActiveSheet.Unprotect

    On Error GoTo 0

End Sub

