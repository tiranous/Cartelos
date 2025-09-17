VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
    On Error GoTo EH
    Engine_RunDefault
    Exit Sub
EH:
    Errors_HandleUnexpected "Sheet1.CommandButton1_Click"
End Sub
