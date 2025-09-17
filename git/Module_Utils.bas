Attribute VB_Name = "Module_Utils"
Option Explicit

Public Type ApplicationStateSnapshot
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    DisplayAlerts As Boolean
End Type

Public Function Utils_DisableForProcessing() As ApplicationStateSnapshot
    Dim snapshot As ApplicationStateSnapshot
    snapshot = Utils_CaptureApplicationState()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With

    Utils_DisableForProcessing = snapshot
End Function

Public Function Utils_CaptureApplicationState() As ApplicationStateSnapshot
    Dim snapshot As ApplicationStateSnapshot

    With Application
        snapshot.ScreenUpdating = .ScreenUpdating
        snapshot.EnableEvents = .EnableEvents
        snapshot.Calculation = .Calculation
        snapshot.DisplayAlerts = .DisplayAlerts
    End With

    Utils_CaptureApplicationState = snapshot
End Function

Public Sub Utils_RestoreApplicationState(ByVal snapshot As ApplicationStateSnapshot)
    On Error Resume Next
    With Application
        .ScreenUpdating = snapshot.ScreenUpdating
        .EnableEvents = snapshot.EnableEvents
        .Calculation = snapshot.Calculation
        .DisplayAlerts = snapshot.DisplayAlerts
    End With
    On Error GoTo 0
End Sub

Public Function Utils_IsEmptyValue(ByVal candidate As Variant) As Boolean
    Utils_IsEmptyValue = (VarType(candidate) = vbEmpty) Or _
                         (VarType(candidate) = vbNull) Or _
                         (Trim$(CStr(candidate)) = vbNullString)
End Function

Public Function Utils_NullSafeString(ByVal candidate As Variant) As String
    If Utils_IsEmptyValue(candidate) Then
        Utils_NullSafeString = vbNullString
    Else
        Utils_NullSafeString = CStr(candidate)
    End If
End Function

Public Function Utils_NormalizePath(ByVal folder As String) As String
    Dim trimmed As String
    trimmed = Trim$(folder)

    If trimmed = vbNullString Then
        Utils_NormalizePath = vbNullString
        Exit Function
    End If

    If Right$(trimmed, 1) <> Application.PathSeparator Then
        Utils_NormalizePath = trimmed & Application.PathSeparator
    Else
        Utils_NormalizePath = trimmed
    End If
End Function

Public Function Utils_ToSafeFileName(ByVal candidate As String) As String
    Dim invalidChars As Variant
    invalidChars = Array("\\", "/", ":", "*", "?", "\"", "<", ">", "|")

    Dim result As String
    result = candidate

    Dim index As Long
    For index = LBound(invalidChars) To UBound(invalidChars)
        result = Replace$(result, invalidChars(index), "_")
    Next index

    Utils_ToSafeFileName = Trim$(result)
End Function

Public Function Utils_ArrayContains(ByVal items As Variant, ByVal valueToFind As String) As Boolean
    Dim index As Long
    For index = LBound(items) To UBound(items)
        If StrComp(CStr(items(index)), valueToFind, vbTextCompare) = 0 Then
            Utils_ArrayContains = True
            Exit Function
        End If
    Next index
    Utils_ArrayContains = False
End Function

Public Function Utils_CombinePath(ByVal folder As String, ByVal fileName As String) As String
    If folder = vbNullString Then
        Utils_CombinePath = fileName
    Else
        Utils_CombinePath = Utils_NormalizePath(folder) & fileName
    End If
End Function
