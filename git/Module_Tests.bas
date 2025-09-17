Attribute VB_Name = "Module_Tests"
Option Explicit

Public Sub RunAllTests()
    On Error GoTo EH

    Dim testsWs As Worksheet
    Set testsWs = IO_EnsureWorksheet("Tests")
    testsWs.Cells.ClearContents

    Logging_Info "Εκτέλεση δοκιμών", "Tests"

    Tests_RecordResult "Έλεγχος ρυθμίσεων", Tests_CheckConfig(), "Η προεπιλεγμένη διαμόρφωση είναι έγκυρη"
    Tests_RecordResult "Έλεγχος logging", Tests_CheckLogging(), "Γράφτηκε γραμμή στο φύλλο Logs"
    Tests_RecordResult "Δοκιμαστική εκτέλεση", Tests_RunEngineDryRun(), "Εκτελέστηκε ο κινητήρας χωρίς αποθήκευση"

    Logging_Info "Ολοκλήρωση δοκιμών", "Tests"
    Exit Sub

EH:
    Errors_HandleUnexpected "Module_Tests.RunAllTests"
End Sub

Private Sub Tests_RecordResult(ByVal name As String, ByVal passed As Boolean, ByVal details As String)
    IO_WriteTestResult name, passed, details
End Sub

Private Function Tests_CheckConfig() As Boolean
    On Error GoTo Fail
    Dim config As CardGenerationConfig
    config = Engine_GetDefaultConfig()
    Errors_ValidateConfig config
    Tests_CheckConfig = True
    Exit Function
Fail:
    Logging_Error "Tests_CheckConfig", Err.Description
    Tests_CheckConfig = False
End Function

Private Function Tests_CheckLogging() As Boolean
    On Error GoTo Fail
    Logging_Info "Δοκιμαστική εγγραφή", "Tests_CheckLogging"
    Tests_CheckLogging = True
    Exit Function
Fail:
    Tests_CheckLogging = False
End Function

Private Function Tests_RunEngineDryRun() As Boolean
    On Error GoTo Fail
    Dim config As CardGenerationConfig
    config = Engine_GetDefaultConfig()
    config.SaveEnabled = False
    config.EndEmployeeRow = config.StartEmployeeRow
    Engine_RunWithConfig config
    Tests_RunEngineDryRun = True
    Exit Function
Fail:
    Logging_Error "Tests_RunEngineDryRun", Err.Description
    Tests_RunEngineDryRun = False
End Function
