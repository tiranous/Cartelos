Attribute VB_Name = "Module_Engine"
Option Explicit

Public Type CardGenerationConfig
    SourceSheetName As String
    TemplateSheetName As String
    OutputSheetName As String
    LookupSheetName As String
    SettingsSheetName As String
    StartEmployeeRow As Long
    EndEmployeeRow As Long
    ScheduleStartRow As Long
    DaysToProcess As Long
    ScheduleFirstColumn As Long
    ScheduleLastColumn As Long
    StartScheduleColumn As Long
    LookupStartRow As Long
    LookupEndRow As Long
    SourceEmployeeColumn As Long
    LookupEmployeeColumn As Long
    LookupBadgeColumn As Long
    LookupContractColumn As Long
    EmployeeCodeTargetAddress As String
    TemplateRange As String
    SaveFolder As String
    SaveSuffix As String
    SaveExtension As String
    SaveEnabled As Boolean
    SpecialBadgesWithoutContract() As String
End Type

Private Type EmployeeLookupResult
    Found As Boolean
    CodeValue As String
    BadgeValue As String
    ContractValue As String
End Type

Public Sub Engine_RunDefault()
    Dim config As CardGenerationConfig
    config = Engine_GetDefaultConfig()
    Engine_RunWithConfig config
End Sub

Public Sub Engine_RunSingle(ByVal employeeRow As Long, Optional ByVal saveFiles As Boolean = False)
    Dim config As CardGenerationConfig
    config = Engine_GetDefaultConfig()
    config.StartEmployeeRow = employeeRow
    config.EndEmployeeRow = employeeRow
    config.SaveEnabled = saveFiles
    Engine_RunWithConfig config
End Sub

Public Sub Engine_RunWithConfig(ByVal config As CardGenerationConfig)
    On Error GoTo EH

    Errors_ValidateConfig config

    Dim sourceWs As Worksheet
    Dim templateWs As Worksheet
    Dim outputWs As Worksheet
    Dim lookupWs As Worksheet

    Set sourceWs = IO_GetWorksheet(config.SourceSheetName)
    Set templateWs = IO_GetWorksheet(config.TemplateSheetName)
    Set outputWs = IO_GetWorksheet(config.OutputSheetName)
    Set lookupWs = IO_GetWorksheet(config.LookupSheetName)

    Dim state As ApplicationStateSnapshot
    Dim stateCaptured As Boolean
    Dim hadError As Boolean
    state = Utils_DisableForProcessing()
    stateCaptured = True

    Logging_Info "Έναρξη παραγωγής καρτελών", "Engine"

    Dim employeeRow As Long
    For employeeRow = config.StartEmployeeRow To config.EndEmployeeRow
        Engine_ProcessEmployee sourceWs, templateWs, outputWs, lookupWs, config, employeeRow
    Next employeeRow

Done:
    If stateCaptured Then
        Utils_RestoreApplicationState state
    End If

    If hadError Then
        Logging_Warning "Η διαδικασία ολοκληρώθηκε με σφάλματα", "Engine"
    Else
        Logging_Info "Ολοκλήρωση παραγωγής καρτελών", "Engine"
    End If
    Exit Sub

EH:
    hadError = True
    Logging_Error "Engine_RunWithConfig", "#" & Err.Number & " " & Err.Description
    Resume Done
End Sub

Public Function Engine_GetDefaultConfig() As CardGenerationConfig
    Dim config As CardGenerationConfig

    config.SourceSheetName = "Sheet1"
    config.TemplateSheetName = "Sheet3"
    config.OutputSheetName = "Sheet2"
    config.LookupSheetName = "Sheet4"
    config.SettingsSheetName = "Sheet5"

    config.StartEmployeeRow = 15
    config.EndEmployeeRow = 25
    config.ScheduleStartRow = 13
    config.DaysToProcess = 28
    config.ScheduleFirstColumn = 3
    config.ScheduleLastColumn = 8
    config.StartScheduleColumn = 3
    config.LookupStartRow = 2
    config.LookupEndRow = 80
    config.SourceEmployeeColumn = 2
    config.LookupEmployeeColumn = 3
    config.LookupBadgeColumn = 2
    config.LookupContractColumn = 4
    config.EmployeeCodeTargetAddress = "G45"
    config.TemplateRange = "A1:O60"
    config.SaveFolder = Utils_NormalizePath(ThisWorkbook.Path)
    config.SaveSuffix = ""
    config.SaveExtension = ".xls"
    config.SaveEnabled = True

    Dim badges As Variant
    badges = Array("90087332", "90087495")
    config.SpecialBadgesWithoutContract = badges

    Dim settingsWs As Worksheet
    Set settingsWs = IO_TryGetWorksheet(config.SettingsSheetName)

    If Not settingsWs Is Nothing Then
        config.StartEmployeeRow = IO_ReadLongSetting(settingsWs, "B1", config.StartEmployeeRow)
        config.EndEmployeeRow = IO_ReadLongSetting(settingsWs, "B2", config.EndEmployeeRow)
        config.DaysToProcess = IO_ReadLongSetting(settingsWs, "B15", config.DaysToProcess)
        config.EmployeeCodeTargetAddress = IO_ReadStringSetting(settingsWs, "B13", config.EmployeeCodeTargetAddress)
        config.SaveFolder = Utils_NormalizePath(IO_ReadStringSetting(settingsWs, "B20", config.SaveFolder))
        config.SaveSuffix = IO_ReadStringSetting(settingsWs, "B22", config.SaveSuffix)
    End If

    Engine_GetDefaultConfig = config
End Function

Private Sub Engine_ProcessEmployee(ByVal sourceWs As Worksheet, _
                                   ByVal templateWs As Worksheet, _
                                   ByVal outputWs As Worksheet, _
                                   ByVal lookupWs As Worksheet, _
                                   ByVal config As CardGenerationConfig, _
                                   ByVal employeeRow As Long)
    Dim employeeCode As String
    employeeCode = Trim$(Utils_NullSafeString(sourceWs.Cells(employeeRow, config.SourceEmployeeColumn).Value))

    If employeeCode = vbNullString Then
        Logging_Warning "Παράλειψη κενής γραμμής", "Row " & CStr(employeeRow)
        Exit Sub
    End If

    IO_ResetOutput templateWs, outputWs, config.TemplateRange

    Dim lookupResult As EmployeeLookupResult
    lookupResult = Engine_FindEmployee(lookupWs, config, employeeCode)

    Engine_PopulateHeader outputWs, config, employeeCode, lookupResult

    Engine_FillSchedule outputWs, sourceWs, config, employeeRow

    Engine_SaveEmployee outputWs, config, employeeCode
End Sub

Private Function Engine_FindEmployee(ByVal lookupWs As Worksheet, _
                                     ByVal config As CardGenerationConfig, _
                                     ByVal employeeCode As String) As EmployeeLookupResult
    Dim result As EmployeeLookupResult

    Dim matchRow As Long
    matchRow = IO_FindMatchRow(lookupWs, config.LookupEmployeeColumn, employeeCode, config.LookupStartRow, config.LookupEndRow)

    If matchRow > 0 Then
        result.Found = True
        result.CodeValue = Utils_NullSafeString(lookupWs.Cells(matchRow, config.LookupEmployeeColumn).Value)
        result.BadgeValue = Utils_NullSafeString(lookupWs.Cells(matchRow, config.LookupBadgeColumn).Value)
        result.ContractValue = Utils_NullSafeString(lookupWs.Cells(matchRow, config.LookupContractColumn).Value)
    End If

    Engine_FindEmployee = result
End Function

Private Sub Engine_PopulateHeader(ByVal outputWs As Worksheet, _
                                  ByVal config As CardGenerationConfig, _
                                  ByVal employeeCode As String, _
                                  ByVal lookupResult As EmployeeLookupResult)
    Dim displayValue As String
    Dim contractValue As String
    Dim badgeValue As String

    If lookupResult.Found Then
        badgeValue = Trim$(lookupResult.BadgeValue)
        displayValue = lookupResult.CodeValue & " (" & badgeValue & ")"
        contractValue = lookupResult.ContractValue
        outputWs.Range(config.EmployeeCodeTargetAddress).Value = lookupResult.CodeValue
    Else
        displayValue = employeeCode
        contractValue = vbNullString
        badgeValue = vbNullString
        outputWs.Range(config.EmployeeCodeTargetAddress).Value = employeeCode
        Logging_Warning "Δεν βρέθηκε αντιστοίχιση στο Sheet4", employeeCode
    End If

    outputWs.Range("C6").Value = displayValue
    outputWs.Range("C8").Value = contractValue

    If badgeValue <> vbNullString Then
        If Utils_ArrayContains(config.SpecialBadgesWithoutContract, badgeValue) Then
            outputWs.Range("C8").Value = vbNullString
        End If
    End If

    If Utils_IsEmptyValue(outputWs.Range("C8").Value) Then
        IO_WriteDefaultShiftSummary outputWs
    End If
End Sub

Private Sub Engine_FillSchedule(ByVal outputWs As Worksheet, _
                                ByVal sourceWs As Worksheet, _
                                ByVal config As CardGenerationConfig, _
                                ByVal employeeRow As Long)
    Dim dayIndex As Long
    For dayIndex = 0 To config.DaysToProcess - 1
        Dim sourceColumn As Long
        sourceColumn = config.StartScheduleColumn + dayIndex

        Dim sourceValue As Variant
        sourceValue = sourceWs.Cells(employeeRow, sourceColumn).Value

        Dim targetRow As Long
        targetRow = config.ScheduleStartRow + dayIndex

        If Utils_IsEmptyValue(sourceValue) Then
            If Utils_IsEmptyValue(outputWs.Range("C8").Value) Then
                IO_WriteShiftTimes outputWs, targetRow, config.ScheduleFirstColumn, config.ScheduleFirstColumn + 1, "07:00", "15:00"
                IO_WriteShiftTimes outputWs, targetRow, config.ScheduleFirstColumn + 2, config.ScheduleFirstColumn + 3, "15:00", "23:00"
                IO_WriteShiftTimes outputWs, targetRow, config.ScheduleFirstColumn + 4, config.ScheduleLastColumn, "23:00", "07:00"
            End If
            IO_WriteShiftRow outputWs, targetRow, config.ScheduleFirstColumn, config.ScheduleLastColumn
        End If
    Next dayIndex
End Sub

Private Sub Engine_SaveEmployee(ByVal outputWs As Worksheet, _
                                ByVal config As CardGenerationConfig, _
                                ByVal employeeCode As String)
    If Not config.SaveEnabled Then
        Exit Sub
    End If

    Dim filePath As String
    filePath = Engine_BuildFilePath(config, employeeCode)

    If filePath = vbNullString Then
        Logging_Warning "Δεν ήταν δυνατή η αποθήκευση - άδειος φάκελος", employeeCode
        Exit Sub
    End If

    On Error GoTo SaveError
    IO_SaveOutputWorksheet outputWs, filePath
    Logging_Info "Αποθήκευση καρτέλας", filePath
    Exit Sub

SaveError:
    Logging_Error "Engine_SaveEmployee", "Αποτυχία αποθήκευσης " & filePath & ": " & Err.Description
End Sub

Private Function Engine_BuildFilePath(ByVal config As CardGenerationConfig, _
                                      ByVal employeeCode As String) As String
    Dim baseFolder As String
    baseFolder = Utils_NormalizePath(config.SaveFolder)

    If baseFolder = vbNullString Then
        baseFolder = Utils_NormalizePath(ThisWorkbook.Path)
    End If

    If baseFolder = vbNullString Then
        Engine_BuildFilePath = vbNullString
        Exit Function
    End If

    Dim safeName As String
    safeName = Utils_ToSafeFileName(employeeCode)
    If safeName = vbNullString Then
        safeName = "card"
    End If

    Dim extension As String
    extension = config.SaveExtension
    If extension = vbNullString Then
        extension = ".xls"
    ElseIf Left$(extension, 1) <> "." Then
        extension = "." & extension
    End If

    Engine_BuildFilePath = baseFolder & safeName & config.SaveSuffix & extension
End Function
