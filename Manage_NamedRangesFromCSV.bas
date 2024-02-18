Attribute VB_Name = "Manage_NamedRangesFromCSV"
'
'Public
'

Sub ExportNamedRangesToCSV()
    Dim fileName As String
    Dim rangeList As Collection
    
    ' Get the collection of named range-cell pairs
    Set rangeList = GetNamedRangesList(ThisWorkbook)
    
    ' Prompt the user for the CSV file name
    fileName = GetSaveCSVFileName
    
    ' If no file selected, exit the subroutine
    If fileName = "" Then Exit Sub
    
    ' Write named range-cell pairs to the CSV file
    WriteRangeListToCSV rangeList, fileName
    
    MsgBox "Named ranges exported to " & fileName, vbInformation
End Sub

Sub ImportNamedRangesFromCSV()
    Dim fileName As String
    Dim rangeList As Collection
    
    ' Prompt the user to select the CSV file
    fileName = GetOpenCSVFileName
    
    ' If no file selected, exit the subroutine
    If fileName = "" Then Exit Sub
    
    ' Read named range-cell pairs from the CSV file
    Set rangeList = ReadRangeListFromCSV(fileName)
    
    ' Clear existing named ranges
    ClearNamedRanges ThisWorkbook
    
    ' Add named ranges to the workbook
    AddNamedRanges ThisWorkbook, rangeList
    
    MsgBox "Named ranges imported from " & fileName, vbInformation
End Sub

Sub UpdateNamedRangesFromCSV()
    Dim fileName As String
    Dim rangeList As Collection
    
    ' Prompt the user to select the CSV file
    fileName = GetOpenCSVFileName
    
    ' If no file selected, exit the subroutine
    If fileName = "" Then Exit Sub
    
    ' Read named range-cell pairs from the CSV file
    Set rangeList = ReadRangeListFromCSV(fileName)
    
    ' Update named ranges in the workbook
    UpdateNamedRanges ThisWorkbook, rangeList
    
    MsgBox "Named ranges updated from " & fileName, vbInformation
End Sub


'
'File
'

Private Function GetSaveCSVFileName() As String
    Dim fileName As String
    
    ' Prompt the user to select the CSV file
    fileName = GetCSVFileName("Save As CSV File", "NamedRanges.csv")
    
    ' If no file selected, display a message and return an empty string
    If fileName = "" Then MsgBox "No file selected. Exiting.", vbExclamation
    GetSaveCSVFileName = fileName
End Function


Private Function GetOpenCSVFileName() As String
    Dim fileName As String
    
    ' Prompt the user to select the CSV file
    fileName = GetCSVFileName("Open CSV File", "NamedRanges.csv")
    
    ' If no file selected, display a message and return an empty string
    If fileName = "" Then MsgBox "No file selected. Exiting.", vbExclamation
    GetOpenCSVFileName = fileName
End Function


Private Function GetCSVFileName(ByVal dialogTitle As String, ByVal defaultFileName As String) As String
    Dim fileName As Variant
    
    ' Set the initial directory to the current folder
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path
    
    ' Prompt the user to select the CSV file
    fileName = Application.GetSaveAsFilename(FileFilter:="CSV Files (*.csv), *.csv", _
                                             Title:=dialogTitle, _
                                             InitialFileName:=defaultFileName)

        ' Check if user canceled the operation or no file selected
    If VarType(fileName) <> vbBoolean Then GetCSVFileName = CStr(fileName)
    
End Function


'
'Collection
'

Private Function GetNamedRangesList(ByVal wb As Workbook) As Collection
    Dim nm As Name
    Dim rangeList As New Collection ' Collection to store named range-cell pairs
    
    ' Loop through all named ranges in the workbook
    For Each nm In wb.Names
        ' Check if named range is special and should be skipped
        If Not IsSpecialNamedRange(nm) Then
            ' Check if named range refers to a range on a worksheet
            If Not nm.RefersToRange Is Nothing Then
                ' Store named range name and its cell reference in the list
                rangeList.Add Array(nm.Name, nm.RefersToRange.Worksheet.Name & "!" & nm.RefersToRange.Address)
            End If
        End If
    Next nm
    
    ' Return the collection of named range-cell pairs
    Set GetNamedRangesList = rangeList
End Function

Private Function ReadRangeListFromCSV(ByVal fileName As String) As Collection
    Dim rangeList As New Collection
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    
    ' Open the CSV file for reading
    fileNum = FreeFile()
    Open fileName For Input As fileNum
    
    ' Read headers from the first line (discard it)
    Line Input #fileNum, line
    
    ' Read data from subsequent lines and add named ranges
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, ",")
        rangeList.Add Array(parts(0), parts(1))
    Loop
    
    ' Close the CSV file
    Close fileNum
    
    Set ReadRangeListFromCSV = rangeList
End Function

Private Sub WriteRangeListToCSV(ByVal rangeList As Collection, ByVal fileName As String)
    Dim fileNum As Integer
    Dim pair
    
    ' Open the file for writing
    fileNum = FreeFile()
    Open fileName For Output As fileNum
    
    ' Write headers to CSV file
    Print #fileNum, "Named Range,Cell Reference"
    
    ' Loop through the list of named range-cell pairs and write them to the CSV file
    For Each pair In rangeList
        Print #fileNum, pair(0) & "," & pair(1)
    Next pair
    
    ' Close the file
    Close fileNum
End Sub

'
'Range
'

Private Sub ClearNamedRanges(ByVal wb As Workbook)
    Dim n As Long
    
    ' Clear existing named ranges
    For n = wb.Names.Count To 1 Step -1
        ' Check if the named range should be skipped
        If Not IsSpecialNamedRange(wb.Names(n)) Then
            wb.Names(n).Delete
        End If
    Next n
End Sub

Private Sub AddNamedRanges(ByVal wb As Workbook, ByVal rangeList As Collection)
    Dim pair
    
    ' Loop through the collection of named range-cell pairs and add them to the workbook
    For Each pair In rangeList
        AddNamedRange wb, pair(0), pair(1)
    Next pair
End Sub

Private Sub AddNamedRange(ByVal wb As Workbook, ByVal namedRange As String, ByVal cellRef As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim sheetName As String
    Dim cellAddress As String
    Dim parts() As String
    
    ' Extract sheet name and cell address from cellRef
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

Private Sub UpdateNamedRanges(ByVal wb As Workbook, ByVal rangeList As Collection)
    Dim pair
    
    ' Loop through the collection of named range-cell pairs and update the named ranges in the workbook
    For Each pair In rangeList
        UpdateNamedRange wb, pair(0), pair(1)
    Next pair
End Sub

Private Sub UpdateNamedRange(ByVal wb As Workbook, ByVal namedRange As String, ByVal cellRef As String)
    Dim namedRangeObj As Name
    Dim newRef As Range
    
    ' Check if the named range exists in the workbook
    On Error Resume Next
    Set namedRangeObj = wb.Names(namedRange)
    On Error GoTo 0
    
    If namedRangeObj Is Nothing Then
        MsgBox "Named range '" & namedRange & "' not found. Skipping update.", vbExclamation
        Exit Sub
    End If
    
    ' Set the new range reference
    Set newRef = wb.Sheets(Split(cellRef, "!")(0)).Range(Split(cellRef, "!")(1))
    
    ' Update the named range reference
    namedRangeObj.RefersTo = newRef
End Sub

Private Function IsSpecialNamedRange(ByVal nm As Name) As Boolean
    ' Check if the named range should be skipped based on its name
    ' Add more conditions as needed to skip other special named ranges
    If Left(nm.Name, 6) = "_xlfn." Then
        IsSpecialNamedRange = True
    Else
        IsSpecialNamedRange = False
    End If
End Function

