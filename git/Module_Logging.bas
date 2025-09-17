Attribute VB_Name = "Module_Logging"
Option Explicit

Private Const LOG_SHEET_NAME As String = "Logs"

Public Sub Logging_Info(ByVal message As String, Optional ByVal context As String = "")
    Logging_Write "INFO", message, context
End Sub

Public Sub Logging_Warning(ByVal message As String, Optional ByVal context As String = "")
    Logging_Write "WARN", message, context
End Sub

Public Sub Logging_Error(ByVal source As String, ByVal message As String)
    Logging_Write "ERROR", message, source
End Sub

Private Sub Logging_Write(ByVal level As String, ByVal message As String, ByVal context As String)
    Dim logSheet As Worksheet
    Set logSheet = IO_EnsureWorksheet(LOG_SHEET_NAME)

    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    logSheet.Cells(1, 1).Value = "Ημερομηνία"
    logSheet.Cells(1, 2).Value = "Επίπεδο"
    logSheet.Cells(1, 3).Value = "Μήνυμα"
    logSheet.Cells(1, 4).Value = "Πλαίσιο"

    logSheet.Cells(nextRow, 1).Value = Now
    logSheet.Cells(nextRow, 2).Value = level
    logSheet.Cells(nextRow, 3).Value = message
    logSheet.Cells(nextRow, 4).Value = context
End Sub
