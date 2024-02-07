Sub ExportNamedRangesToCSV()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim nm As Name
    Dim fileName As String
    Dim fileNum As Integer
    
    ' Set the workbook object
    Set wb = ThisWorkbook
    
    ' Prompt the user for the CSV file name
    fileName = GetSaveCSVFileName
    
    
    ' Open the file for writing
    fileNum = FreeFile()
    Open fileName For Output As fileNum
    
    ' Write headers to CSV file
    Print #fileNum, "Named Range,Cell Reference"
    
    ' Loop through all named ranges in the workbook
    For Each nm In wb.Names
        ' Check if named range refers to a range on a worksheet
        If Not nm.RefersToRange Is Nothing Then
            ' Write named range name and its cell reference to CSV file
            Print #fileNum, nm.Name & "," & nm.RefersToRange.Worksheet.Name & "!" & nm.RefersToRange.Address
        End If
    Next nm
    
    ' Close the file
    Close fileNum
    
    MsgBox "Named ranges exported to " & fileName, vbInformation
End Sub


Private Function GetSaveCSVFileName() As String
    Dim fileName As String
    
    ' Prompt the user to select the CSV file
    fileName = Application.GetSaveAsFilename(FileFilter:="CSV Files (*.csv), *.csv")
    
    ' Check if user canceled the operation
    If fileName = "False" Then
        GetSaveCSVFileName = ""
    Else
        GetSaveCSVFileName = fileName
    End If
End Function



Sub ImportNamedRangesFromCSV()
    Dim wb As Workbook
    Dim fileName As String
    
    ' Set the workbook object
    Set wb = ThisWorkbook
    
    ' Prompt the user to select the CSV file
    fileName = GetOpenCSVFileName
    
    ' Perform the import process
    ' Open the CSV file for reading
    fileNum = FreeFile()
    Open fileName For Input As fileNum
    
    ' Clear existing named ranges
    ClearNamedRanges wb
    
    ' Read headers from the first line (discard it)
    Line Input #fileNum, line
    
    ' Read data from subsequent lines and add named ranges
    ReadAndAddNamedRanges wb, fileNum
    
    ' Close the CSV file
    Close fileNum
    
    MsgBox "Named ranges imported from " & fileName, vbInformation
End Sub

Private Function GetOpenCSVFileName() As String
    Dim fileName As String
    
    ' Prompt the user to select the CSV file
    fileName = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV File")
    
    ' Check if user canceled the operation
    If fileName = "False" Then
        GetOpenCSVFileName = ""
    Else
        GetOpenCSVFileName = fileName
    End If
End Function


Private Sub ReadAndAddNamedRanges(wb As Workbook, fileNum As Variant)
    Dim line As String
    Dim parts() As String
    Dim namedRange As String
    Dim cellRef As String
    
    ' Read data from subsequent lines and add named ranges
    Do Until EOF(fileNum)
        ' Read each line
        Line Input #fileNum, line
        ' Split the line by comma delimiter
        parts = Split(line, ",")
        ' Extract named range name and cell reference
        namedRange = parts(0)
        cellRef = parts(1)
        ' Add named range to the workbook
        AddNamedRange wb, namedRange, cellRef
    Loop
End Sub

Private Sub AddNamedRange(wb As Workbook, namedRange As String, cellRef As String)
    Dim ws As Worksheet
    Dim rng As Range
    
    ' Extract sheet name and cell address from cellRef
    Dim sheetName As String
    Dim cellAddress As String
    Dim parts() As String
    parts = Split(cellRef, "!")
    sheetName = parts(0)
    cellAddress = parts(1)
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found. Named range '" & namedRange & "' not added.", vbExclamation
        Exit Sub
    End If
    
    ' Set the range object
    Set rng = ws.Range(cellAddress)
    
    ' Add named range to the workbook
    wb.Names.Add Name:=namedRange, RefersTo:=rng
End Sub

Private Sub ClearNamedRanges(wb As Workbook)
    Dim n As Long
    
    ' Clear existing named ranges
    For n = wb.Names.Count To 1 Step -1
        wb.Names(n).Delete
    Next n
End Sub

