Attribute VB_Name = "Module_IO"
Option Explicit

Public Function IO_GetWorksheet(ByVal sheetName As String) As Worksheet
    Dim target As Worksheet
    On Error Resume Next
    Set target = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If target Is Nothing Then
        Errors_RaiseMissingSheet sheetName
    End If

    Set IO_GetWorksheet = target
End Function

Public Function IO_TryGetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set IO_TryGetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Public Function IO_EnsureWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = IO_TryGetWorksheet(sheetName)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If

    Set IO_EnsureWorksheet = ws
End Function

Public Sub IO_ResetOutput(ByVal templateWs As Worksheet, _
                           ByVal outputWs As Worksheet, _
                           ByVal templateRange As String)
    outputWs.Cells.Clear
    templateWs.Range(templateRange).Copy Destination:=outputWs.Range(templateRange)
End Sub

Public Sub IO_WriteDefaultShiftSummary(ByVal outputWs As Worksheet)
    outputWs.Range("C11").Value = "'       Α΄ ΒΑΡΔΙΑ"
    outputWs.Range("E11").Value = "'       Β΄ ΒΑΡΔΙΑ"
    outputWs.Range("G11").Value = "'       Γ΄ ΒΑΡΔΙΑ"
End Sub

Public Sub IO_WriteShiftRow(ByVal outputWs As Worksheet, _
                             ByVal rowIndex As Long, _
                             ByVal firstColumn As Long, _
                             ByVal lastColumn As Long)
    Dim target As Range
    Set target = outputWs.Range(outputWs.Cells(rowIndex, firstColumn), outputWs.Cells(rowIndex, lastColumn))

    With target
        .UnMerge
        .Value = " "
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Merge
    End With
End Sub

Public Sub IO_WriteShiftTimes(ByVal outputWs As Worksheet, _
                              ByVal rowIndex As Long, _
                              ByVal startColumn As Long, _
                              ByVal endColumn As Long, _
                              ByVal startValue As String, _
                              ByVal endValue As String)
    outputWs.Cells(rowIndex, startColumn).Value = startValue
    outputWs.Cells(rowIndex, endColumn).Value = endValue
End Sub

Public Sub IO_SaveOutputWorksheet(ByVal outputWs As Worksheet, ByVal filePath As String)
    Dim newBook As Workbook

    outputWs.Copy
    Set newBook = ActiveWorkbook

    On Error GoTo CleanUp
    newBook.SaveAs Filename:=filePath, FileFormat:=xlWorkbookNormal
    newBook.Close SaveChanges:=False
    Exit Sub

CleanUp:
    On Error Resume Next
    newBook.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function IO_ReadLongSetting(ByVal settingsWs As Worksheet, _
                                   ByVal address As String, _
                                   ByVal defaultValue As Long) As Long
    If settingsWs Is Nothing Then
        IO_ReadLongSetting = defaultValue
        Exit Function
    End If

    Dim candidate As Variant
    candidate = settingsWs.Range(address).Value

    If IsNumeric(candidate) Then
        IO_ReadLongSetting = CLng(candidate)
    Else
        IO_ReadLongSetting = defaultValue
    End If
End Function

Public Function IO_ReadStringSetting(ByVal settingsWs As Worksheet, _
                                     ByVal address As String, _
                                     ByVal defaultValue As String) As String
    If settingsWs Is Nothing Then
        IO_ReadStringSetting = defaultValue
        Exit Function
    End If

    Dim candidate As Variant
    candidate = settingsWs.Range(address).Value

    Dim valueText As String
    valueText = Utils_NullSafeString(candidate)

    If valueText = vbNullString Then
        IO_ReadStringSetting = defaultValue
    Else
        IO_ReadStringSetting = valueText
    End If
End Function

Public Function IO_FindMatchRow(ByVal sourceWs As Worksheet, _
                                ByVal columnIndex As Long, _
                                ByVal valueToFind As String, _
                                ByVal firstRow As Long, _
                                ByVal lastRow As Long) As Long
    Dim rowIndex As Long
    For rowIndex = firstRow To lastRow
        If StrComp(Utils_NullSafeString(sourceWs.Cells(rowIndex, columnIndex).Value), valueToFind, vbTextCompare) = 0 Then
            IO_FindMatchRow = rowIndex
            Exit Function
        End If
    Next rowIndex
    IO_FindMatchRow = 0
End Function

Public Sub IO_WriteTestResult(ByVal testName As String, ByVal passed As Boolean, ByVal details As String)
    Dim testsWs As Worksheet
    Set testsWs = IO_EnsureWorksheet("Tests")

    Dim nextRow As Long
    nextRow = testsWs.Cells(testsWs.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    testsWs.Cells(1, 1).Value = "Test"
    testsWs.Cells(1, 2).Value = "Result"
    testsWs.Cells(1, 3).Value = "Details"

    testsWs.Cells(nextRow, 1).Value = testName
    testsWs.Cells(nextRow, 2).Value = IIf(passed, "Pass", "Fail")
    testsWs.Cells(nextRow, 3).Value = details
End Sub
