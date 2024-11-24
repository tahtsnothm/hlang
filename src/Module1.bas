Attribute VB_Name = "Module1"
Sub tt()
Dim fso As New FileSystemObject
Dim ts As TextStream
Set ts = fso.OpenTextFile("D:\test.txt")

Dim wodecode As String
wodecode = ts.ReadAll
wodecode = Replace(wodecode, vbCrLf, vbLf)

Dim i As Long
For i = 0 To Len(wodecode) - 1
    pt Asc(Mid(wodecode, i + 1, 1)), Mid(wodecode, i + 1, 1)
Next i
ts.Close
End Sub

Sub HI()
Dim hh As New H

hh.runFile "C:\Users\aweso\OneDrive\Desktop\hlang\test.txt"
End Sub


Public Sub Export()

  Dim wbPath As String
  Dim vbComp As Object
  Dim exportPath As String
  wbPath = ThisWorkbook.path & "\build"

  For Each vbComp In ActiveWorkbook.VBProject.VBComponents
    If (vbComp.name = "ThisWorkbook") Then GoTo endloop
    exportPath = wbPath & "\" & vbComp.name '& format$(Now, "_yyyymmdd_hhnnss")
    

    Select Case vbComp.Type
        Case vbext_ct_StdModule ' Standard Module
            exportPath = exportPath & ".bas"
        'Case 2 ' UserForm
           ' exportPath = exportPath & ".frm"
        Case vbext_ct_ClassModule ' Class Module
            exportPath = exportPath & ".cls"
        Case Else ' Anything else
            GoTo endloop
            exportPath = exportPath & ".bas"
    End Select

    On Error Resume Next
    vbComp.Export exportPath
    On Error GoTo 0
endloop:
  Next

End Sub
