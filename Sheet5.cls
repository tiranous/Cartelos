VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Dim i As Integer
Dim j As Integer
Dim k As Long
Dim a As Integer
Dim b As Integer
Dim r As Integer
Dim StartRange As String
Dim EndRange As String
Dim t As Integer

Application.DisplayAlerts = False
Application.CutCopyMode = False



r = 13 'Seira stin kartela
'b = "C" 'ksekinima sto programma
a = 3
'b = 15
j = 15 'ksekinima sto programma
StartRange = "C" & r
EndRange = "H" & r

For b = Sheet5.Range("B1") To Sheet5.Range("B2")
a = 3
j = 15
StartRange = "C" & r
EndRange = "H" & r

For i = 1 To Sheet5.Range("B15") 'loopa mera pros mera

For t = 2 To 80
If Sheet1.Range("B" & b) = Sheet4.Range("C" & t) Then

Sheet2.Range("C6") = Sheet4.Range("C" & t) & " (" & Sheet4.Range("B" & t) & ")"
Sheet2.Range("C8") = Sheet4.Range("D" & t)
x = Sheet5.Range("B13")
Sheet2.Range(x) = Sheet4.Range("C" & t)
'Sheet2.Range("G45") = Sheet4.Range("C" & t)

If Sheet2.Range("C8") = "еиспяайтояас" Then



Sheet2.Range("C11") = "а' баядиа      07:00-15:00"
Sheet2.Range("E11") = "б' баядиа      15:00-23:00"
Sheet2.Range("G11") = "ц' баядиа      23:00-07:00"

'If Sheet2.Range("C6") = "йакйатфийос лаяиос (90087332)" Then 'gia mario kai karagiorgi

'Sheet2.Range("C11") = "а' баядиа      06:30-14:30"
'Sheet2.Range("E11") = "б' баядиа      14:30-22:30"
'Sheet2.Range("G11") = "ц' баядиа      22:30-06:30"

'End If

'If Sheet2.Range("C6") = "йаяациыяцг цеыяциа (90087495)" Then 'gia mario kai karagiorgi

'Sheet2.Range("C11") = "а' баядиа      06:30-14:30"
'Sheet2.Range("E11") = "б' баядиа      14:30-22:30"
'Sheet2.Range("G11") = "ц' баядиа      22:30-06:30"

'End If


End If
t = t + 100

End If

Next

'Sheet2.Range("C6") = Sheet1.Cells(b, 2)

If Sheet1.Cells(b, a) = "п" Then 'an vrei prwi tote
If Sheet2.Range("C8") = "епоптгс" Then
Sheet2.Range("C" & r) = "6:30" 'vale enarski
Sheet2.Range("D" & r) = "14:30" 'vale liksi
End If

If Sheet2.Range("C8") = "еиспяайтояас" Then
Sheet2.Range("C" & r) = "7:00" 'vale enarski
Sheet2.Range("D" & r) = "15:00" 'vale liksi
End If
'r = r + 1 'paei pio katw stin kartela

End If
If Sheet1.Cells(b, a) = "а" Then
If Sheet2.Range("C8") = "епоптгс" Then
Sheet2.Range("E" & r) = "14:30" 'vale enarski
Sheet2.Range("F" & r) = "22:30" 'vale liksi
End If
If Sheet2.Range("C8") = "еиспяайтояас" Then
Sheet2.Range("E" & r) = "15:00" 'vale enarski
Sheet2.Range("F" & r) = "23:00" 'vale liksi
End If



End If

If Sheet1.Cells(b, a) = "м" Then
If Sheet2.Range("C8") = "епоптгс" Then
Sheet2.Range("G" & r) = "22:30" 'vale enarski
Sheet2.Range("H" & r) = "6:30" 'vale liksi
End If
If Sheet2.Range("C8") = "еиспяайтояас" Then
Sheet2.Range("G" & r) = "23:00" 'vale enarski
Sheet2.Range("H" & r) = "7:00" 'vale liksi
End If


End If

If Sheet1.Cells(b, a) = "" Then

'Sheet2.Range
Sheet2.Range("C" & r) = "яепо" 'vale enarski

Sheet2.Range(StartRange & ":" & EndRange).Font.Bold = True
Sheet2.Range(StartRange & ":" & EndRange).Merge
Sheet2.Range(StartRange & ":" & EndRange).HorizontalAlignment = xlCenter

'r = r + 1

End If
If Sheet1.Cells(b, a) = "йа" Then

Sheet2.Range("C" & r) = "етгсиа адеиа" 'vale enarski

Sheet2.Range(StartRange & ":" & EndRange).Font.Bold = True
Sheet2.Range(StartRange & ":" & EndRange).Merge
Sheet2.Range(StartRange & ":" & EndRange).HorizontalAlignment = xlCenter


End If
r = r + 1
a = a + 1
StartRange = "C" & r
EndRange = "H" & r
Next

If Sheet2.Range("C6") = "йакйатфийос лаяиос (90087332)" Then ' gia kalkatziko kai karagiorgi
Sheet2.Range("C8") = "еиспяайтояас"
End If

If Sheet2.Range("C6") = "йаяациыяцг цеыяциа (90087495)" Then ' gia kalkatziko kai karagiorgi
Sheet2.Range("C8") = "еиспяайтояас"
End If


Application.ScreenUpdating = False
 'Application.DisplayAlerts = False
Sheet2.Activate
Sheet2.Select
Sheet2.Copy
ActiveSheet.SaveAs Filename:=Sheet5.Range("B20") & Sheet1.Cells(b, 2) & Sheet5.Range("B22") & "xls"
Application.ScreenUpdating = True
ActiveWorkbook.Close
'Selection.Clear



'arxi pou ka8arizei tin kartela sto2
Sheet2.Cells.Clear
Sheets("Sheet2").Select 'sheet that I want to copy and paste (resides in macro workbook)
Sheet3.Range("A1:O60").Select
Selection.Copy
Sheet2.Activate 'destination workbook (no macros)
Sheets("Sheet1").Select 'destination worksheet
Sheet2.Range("A1:O60").Select
ActiveSheet.Paste
'edw teleiwnei o ka8arismos
'Application.DisplayAlerts = True


a = 3
r = 13
t = 2
j = 15

Next
Application.DisplayAlerts = True
Application.CutCopyMode = True
End Sub

Private Sub CommandButton2_Click()

Dim i As Integer
Dim j As Integer
Dim k As Long
Dim a As Integer
Dim b As Integer
Dim r As Integer
Dim StartRange As String
Dim EndRange As String
Dim t As Integer
Dim x As String

r = 13 'Seira stin kartela
a = 3
j = 15 'ksekinima sto programma
StartRange = "C" & r
EndRange = "H" & r

b = Sheet5.Range("B5")
For i = 1 To Sheet5.Range("B15") 'loopa mera pros mera

For t = 2 To 60
If Sheet1.Range("B" & b) = Sheet4.Range("C" & t) Then

Sheet2.Range("C6") = Sheet4.Range("C" & t) & " (" & Sheet4.Range("B" & t) & ")"
Sheet2.Range("C8") = Sheet4.Range("D" & t)
x = Sheet5.Range("B13")
Sheet2.Range(x) = Sheet4.Range("C" & t)

t = t + 1

End If

Next

'Sheet2.Range("C6") = Sheet1.Cells(b, 2)

If Sheet1.Cells(b, a) = "п" Then 'an vrei prwi tote
If Sheet2.Range("C8") = "епоптгс" Then
Sheet2.Range("C" & r) = "6:30" 'vale enarski
Sheet2.Range("D" & r) = "2:30" 'vale liksi
End If

If Sheet2.Range("C8") = "еиспяайтояас" Then
Sheet2.Range("C" & r) = "7:00" 'vale enarski
Sheet2.Range("D" & r) = "3:00" 'vale liksi
End If
'r = r + 1 'paei pio katw stin kartela

End If
If Sheet1.Cells(b, a) = "а" Then
If Sheet2.Range("C8") = "епоптгс" Then
Sheet2.Range("E" & r) = "14:30" 'vale enarski
Sheet2.Range("F" & r) = "22:30" 'vale liksi
End If
If Sheet2.Range("C8") = "еиспяайтояас" Then
Sheet2.Range("E" & r) = "15:00" 'vale enarski
Sheet2.Range("F" & r) = "23:00" 'vale liksi
End If



End If

If Sheet1.Cells(b, a) = "м" Then
If Sheet2.Range("C8") = "епоптгс" Then
Sheet2.Range("G" & r) = "22:30" 'vale enarski
Sheet2.Range("H" & r) = "6:30" 'vale liksi
End If
If Sheet2.Range("C8") = "еиспяайтояас" Then
Sheet2.Range("G" & r) = "23:00" 'vale enarski
Sheet2.Range("H" & r) = "7:00" 'vale liksi
End If


End If

If Sheet1.Cells(b, a) = "" Then

Sheet2.Range("C" & r) = "яепо" 'vale enarski

Sheet2.Range(StartRange & ":" & EndRange).Font.Bold = True
Sheet2.Range(StartRange & ":" & EndRange).Merge
Sheet2.Range(StartRange & ":" & EndRange).HorizontalAlignment = xlCenter

'r = r + 1

End If
If Sheet1.Cells(b, a) = "йа" Then

Sheet2.Range("C" & r) = "етгсиа адеиа" 'vale enarski

Sheet2.Range(StartRange & ":" & EndRange).Font.Bold = True
Sheet2.Range(StartRange & ":" & EndRange).Merge
Sheet2.Range(StartRange & ":" & EndRange).HorizontalAlignment = xlCenter


End If
r = r + 1
a = a + 1
StartRange = "C" & r
EndRange = "H" & r
Next

Application.ScreenUpdating = False
 Application.DisplayAlerts = False
 Sheet2.Activate
Sheet2.Select
Sheet2.Copy
ActiveSheet.SaveAs Filename:=Sheet5.Range("B20") & Sheet1.Cells(b, 2) & " " & Sheet5.Range("B22") & "xls"
Application.ScreenUpdating = True
ActiveWorkbook.Close
Application.DisplayAlerts = True

End Sub
