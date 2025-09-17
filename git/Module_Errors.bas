Attribute VB_Name = "Module_Errors"
Option Explicit

Private Const ERR_NAMESPACE As Long = vbObjectError + 512

Public Sub Errors_RaiseMissingSheet(ByVal sheetName As String)
    Err.Raise ERR_NAMESPACE + 1, "Module_Errors", "Δεν βρέθηκε το φύλλο " & sheetName & "."
End Sub

Public Sub Errors_RaiseInvalidConfiguration(ByVal message As String)
    Err.Raise ERR_NAMESPACE + 2, "Module_Errors", message
End Sub

Public Sub Errors_HandleUnexpected(ByVal source As String)
    Dim description As String
    description = Err.Description
    Logging_Error source, description
    MsgBox "Παρουσιάστηκε σφάλμα (" & source & "): " & description, vbCritical + vbOKOnly, "Σφάλμα"
End Sub

Public Sub Errors_ValidateConfig(ByRef config As CardGenerationConfig)
    If config.StartEmployeeRow <= 0 Then
        Errors_RaiseInvalidConfiguration("Μη έγκυρη γραμμή εκκίνησης εργαζομένων.")
    End If

    If config.EndEmployeeRow < config.StartEmployeeRow Then
        Errors_RaiseInvalidConfiguration("Η τελική γραμμή εργαζομένων είναι πριν από την αρχική.")
    End If

    If config.DaysToProcess <= 0 Then
        Errors_RaiseInvalidConfiguration("Ο αριθμός ημερών πρέπει να είναι θετικός.")
    End If

    If config.ScheduleFirstColumn <= 0 Or config.ScheduleLastColumn < config.ScheduleFirstColumn Then
        Errors_RaiseInvalidConfiguration("Μη έγκυρη διάταξη στηλών για το πρόγραμμα.")
    End If

    If config.ScheduleStartRow <= 0 Then
        Errors_RaiseInvalidConfiguration("Μη έγκυρη αρχική γραμμή για το πρόγραμμα.")
    End If
End Sub
