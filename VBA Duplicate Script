Sub HighlightDuplicateRows()
    Dim ws As Worksheet
    Dim rng As Range, rowNum As Range
    Dim dict As Object
    Dim key As String
    Dim lastRow As Long
       
    ' Set the active worksheet
    Set ws = ActiveSheet
       
    ' Prompt user to select range
    On Error Resume Next
    Set rng = Application.InputBox("Select the range to check for duplicate rows:", Type:=8)
    On Error GoTo 0

    ' Exit if no range is selected
    If rng Is Nothing Then Exit Sub

    ' Initialize dictionary to track row uniqueness
    Set dict = CreateObject("Scripting.Dictionary")
       
    ' First Pass: Count occurrences of each row
    Dim rowKeys As Object
    Set rowKeys = CreateObject("Scripting.Dictionary")
    
    For Each rowNum In rng.Rows
        key = ""
        
        ' Concatenate values from all columns in the row
        For i = 1 To rng.Columns.Count
            key = key & "|" & Trim(rowNum.Cells(1, i).Text) ' Using .Text instead of .Value for accuracy
        Next i

        ' Count occurrences
        If rowKeys.exists(key) Then
            rowKeys(key) = rowKeys(key) + 1
        Else
            rowKeys.Add key, 1
        End If
    Next rowNum

    ' Second Pass: Highlight only duplicate rows
    For Each rowNum In rng.Rows
        key = ""
        For i = 1 To rng.Columns.Count
            key = key & "|" & Trim(rowNum.Cells(1, i).Text) ' Using .Text for exact match
        Next i

        ' Highlight only if the row appears more than once
        If rowKeys(key) > 1 Then
            rowNum.Interior.Color = RGB(255, 153, 153) ' Light Red Highlight
        End If
    Next rowNum
       
    MsgBox "Duplicate rows highlighted successfully!", vbInformation
End Sub
