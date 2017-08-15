Attribute VB_Name = "Export_Notepad"
Option Explicit

Sub ExportNotepad()

Dim iCount As Integer
Dim iFile As Integer
Dim lLastRow As Long
Dim rMyRange As Range
Dim rCell As Range
Dim strID As String
Dim vValue As Variant

    Set rMyRange = Application.InputBox("Please select the range of ID's", _
        "Select ID's", Type:=8)

    lLastRow = rMyRange.Rows.Count
    iCount = 0

    For Each rCell In rMyRange
        iCount = iCount + 1
        If iCount = lLastRow Then
            strID = strID & rCell.Value
        'ElseIf Int(iCount / 1) = iCount / 1 Then
        '    strID = strID & rCell.Value & "," & vbNewLine
        Else
            strID = strID & rCell.Value & "," & vbNewLine
            
        End If
    Next rCell
    
    iFile = FreeFile

    Open "c:\temp\" & Application.UserName & "_IDs.txt" For Output As #iFile
    Print #iFile, strID
    Close #iFile
     
    vValue = Shell("notepad.exe c:\temp\" & Application.UserName & "_IDs.txt", 1)
     
    Kill "c:\temp\" & Application.UserName & "_IDs.txt"
    
End Sub


