'*******************************************************************************
' Project:         Macro_ColumnRow_To_NamedRange_InFormulae
' Module:          ColumnRowToNamedRangeInFormulae
' Description:     Macro to convert in all formulae of the selected Range
'                  its ColumnoRow into Named Range if it exists.
'
' Author:          https://github.com/NPueyo
' Created:         2024/02/03
'
' Dependencies:    None
'
'*******************************************************************************
Attribute VB_Name = ColumnRowToNamedRangeInFormulae

Public Sub Macro_ColumnRow_To_NamedRange_InFormulae()
    ' Disable various Excel features to improve performance
    TurnOffOptimizations
    
    Dim ws As Worksheet
    Dim targetRange As Range
    
    ' Loop through all worksheets in the current workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Prompt the user to select a range of cells
        Set targetRange = SelectRange(ws)
        
        ' Apply the defined names to the selected range
        ColumnRowToNamedRangeInFormulae targetRange
    Next ws
    
    ' Re-enable Excel features after the code execution
    TurnOnOptimizations
End Sub

Public Sub ColumnRowToNamedRangeInFormulae(rng As Range)
    ' Apply defined names to the specified range of cells.
    ' Also, replace occurrences of the defined names in the range.

    ' Inputs:
    '   rng: Range - The range of cells to which defined names will be applied.

    Dim Nm As Name
    Dim ref As String
    
    ' Loop through all defined names in the workbook
    For Each Nm In ThisWorkbook.Names
        ' Apply the defined name to the cells in the specified range
        On Error Resume Next
        rng.ApplyNames Names:=Array(Nm.Name)
        On Error GoTo 0
        
        ' Extract the reference of the defined name
        ref = Nm.RefersTo
        ref = Mid(ref, 2)
        
        ' Replace occurrences of the reference with the defined name in the range
        rng.Replace What:=ref, Replacement:=Nm.Name
        
        ' Remove absolute references ($) and replace again
        ref = Replace(ref, "$", "")
        rng.Replace What:=ref, Replacement:=Nm.Name
    Next Nm
End Sub

Private Function SelectRange(ws As Worksheet) As Range
    ' Prompt the user to select a range of cells in the specified worksheet.
    ' Returns the selected range.

    ' Inputs:
    '   ws: Worksheet - The worksheet in which the user is prompted to select a range.
        ' Outputs: Range - Selected range

    On Error Resume Next
    Set SelectRange = Application.InputBox("Select a range of cells in '" & ws.Name & "'", Type:=8)
    On Error GoTo 0
End Function

Private Sub TurnOffOptimizations()
    ' Turn off various Excel features to improve performance

    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .DisplayAlerts = False
    End With
End Sub

Private Sub TurnOnOptimizations()
    ' Turn on various Excel features after the code execution

    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
    End With
End Sub
