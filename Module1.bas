Attribute VB_Name = "Module1"
' Module: ExportAll
Option Explicit

Public Sub ExportAllVBA()
    Dim vb As VBIDE.VBProject
    Dim c As VBIDE.VBComponent
    Dim ext As String
    Dim outDir As String
    
    ' Φάκελος export: ο φάκελος που βρίσκεται το workbook + "\git"
    outDir = ThisWorkbook.Path & "\git"
    
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    
    Set vb = ThisWorkbook.VBProject
    For Each c In vb.VBComponents
        Select Case c.Type
            Case vbext_ct_StdModule: ext = ".bas"
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_MSForm: ext = ".frm"
            Case Else: ext = ".cls"   ' Sheets & ThisWorkbook
        End Select
        
        On Error Resume Next
        c.Export outDir & "\" & c.Name & ext
        On Error GoTo 0
    Next c
    
    MsgBox "Export ολοκληρώθηκε στον φάκελο: " & outDir, vbInformation
End Sub

