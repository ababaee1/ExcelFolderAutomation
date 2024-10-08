Sub CreateFoldersAndLinkForCurrentMonthCorrected()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Specify the sheet if needed
    
    Dim initialRow As Long
    initialRow = 2 ' Assuming data starts from row 2
    
    Dim Y As String
    Y = Format(Now, "MM") ' Current month as a two-digit number
    
    Dim folderPath As String
    folderPath = "C:\Users\" ' Change to your target directory
    
    Dim lastRow As Long, maxExistingX As Integer
    maxExistingX = 0 ' Highest X value for the current month
    
    ' Find the actual last row with data for the current month
    lastRow = initialRow
    Dim row As Long
    For row = ws.Rows.Count To initialRow Step -1
        If ws.Cells(row, 1).Value <> "" And Left(ws.Cells(row, 1).Value, 2) = Y Then
            lastRow = row
            Exit For
        End If
    Next row
    
    ' Determine the highest X value for the current month
    Dim X As Integer ' Declare X here, outside of the For loop
    For row = initialRow To lastRow
        If Left(ws.Cells(row, 1).Value, 2) = Y Then
            X = Val(Mid(ws.Cells(row, 1).Value, 4, Len(ws.Cells(row, 1).Value) - 6))
            If X > maxExistingX Then maxExistingX = X
        End If
    Next row
    
    Dim entriesToGenerate As Integer
    If maxExistingX = 0 Then
        entriesToGenerate = 3
        X = 1
    Else
        entriesToGenerate = 3
        X = maxExistingX + 1
    End If
    
    ' Generate folders and links
    For i = 1 To entriesToGenerate
        Dim folderName As String, cellValue As String
        If X < 10 Then
            folderName = Y & "-00" & X & "-24"
            cellValue = Y & "-00" & X & "/24"
        Else
            folderName = Y & "-0" & X & "-24"
            cellValue = Y & "-0" & X & "/24"
        End If
        
        ' Create the folder if it doesn't exist
        If Len(Dir(folderPath & folderName, vbDirectory)) = 0 Then
            MkDir folderPath & folderName
        End If
        
        ' Find the next empty row to avoid overwriting any cells
        Dim nextEmptyRow As Long
        nextEmptyRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
        
        ' Update cell value and link to the folder
        ws.Cells(nextEmptyRow, 1).Value = cellValue
        ws.Hyperlinks.Add Anchor:=ws.Cells(nextEmptyRow, 1), Address:=folderPath & folderName, TextToDisplay:=cellValue
        
        X = X + 1 ' Increment X within the loop without redeclaring it
    Next i
End Sub

