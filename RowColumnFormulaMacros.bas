'*******************************************************************************
' Module:          RowColumnFormulaMacros
' Description:     Macro to convert range strings into R1C1
'
'
' Author:          https://github.com/NPueyo
' Created:         2024/02/03
'
' Dependencies:    None
'
'*******************************************************************************

Attribute VB_Name = "RowColumnFormulaMacros"

Function RCFormula(ByVal rangeref As String) As String
    ' Split the worksheet name and cell reference
    Dim exclamationPos As Long
    exclamationPos = InStr(rangeref, "!")

    ' Extract the worksheet name and cell reference
    Dim hoja As String
    Dim RC As String

    hoja = Left(rangeref, exclamationPos - 1)
    RC = Mid(rangeref, exclamationPos + 1)

    ' Generate the formula and
    ' Return the generated formula
    RCFormula = hoja & "!" & StringToRC(RC)
End Function



Function StringToRC(cellRef As String) As String
    Dim rowNumber As Long
    Dim colNumber As Long

    ' Convert the cell reference string to uppercase
    cellRef = UCase(cellRef)

    ' Extract the row and column numbers from the cell reference
    rowNumber = Range(cellRef).row
    colNumber = Range(cellRef).Column

    ' Return the formatted string
    StringToRC = "R" & rowNumber & "C" & colNumber
End Function
