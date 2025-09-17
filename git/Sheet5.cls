VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    On Error GoTo EH
    Engine_RunWithConfig Engine_GetDefaultConfig()
    Exit Sub
EH:
    Errors_HandleUnexpected "Sheet5.CommandButton1_Click"
End Sub

Private Sub CommandButton2_Click()
    On Error GoTo EH
    Dim config As CardGenerationConfig
    config = Engine_GetDefaultConfig()
    Dim singleRow As Long
    singleRow = IO_ReadLongSetting(Me, "B5", config.StartEmployeeRow)
    config.StartEmployeeRow = singleRow
    config.EndEmployeeRow = singleRow
    Engine_RunWithConfig config
    Exit Sub
EH:
    Errors_HandleUnexpected "Sheet5.CommandButton2_Click"
End Sub
