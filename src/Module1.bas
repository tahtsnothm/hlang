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


Sub hi()
Dim hh As New H

hh.runFile ThisWorkbook.Path & "\test.txt"

End Sub

Sub hcc(filename As String)
Dim hh As New H

hh.runFile ThisWorkbook.Path & "\" & filename
End Sub


Public Sub Export()

  Dim wbPath As String: wbPath = ThisWorkbook.Path & "\build"
  Dim vbComp As Object

  For Each vbComp In ThisWorkbook.VBProject.VBComponents
    If (vbComp.name = "ThisWorkbook") Then GoTo endloop
    Dim exportPath As String: exportPath = wbPath & "\" & vbComp.name '& format$(Now, "_yyyymmdd_hhnnss")
    
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
    pt "Export:", vbComp.name
    On Error GoTo 0
endloop:
  Next

End Sub

Public Sub import()
' DONT FUCKING TOUCH
Exit Sub
    Export

   Dim MyObj As Object, MySource As Object, file As Variant
   file = Dir(ThisWorkbook.Path & "\src\")
   Dim mset As New Scripting.Dictionary
   Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Dim s As String: s = vbComp.name
        Select Case vbComp.Type
        Case vbext_ct_StdModule ' Standard Module
            s = s & ".bas"
        'Case 2 ' UserForm
           ' exportPath = exportPath & ".frm"
        Case vbext_ct_ClassModule ' Class Module
            s = s & ".cls"
    End Select
        mset.Add s, vbComp
    Next vbComp
   While (file <> "")
        pt "Import:", file
        If (mset.Exists(file)) Then
            ThisWorkbook.VBProject.VBComponents.Remove mset(file)
        End If
        ThisWorkbook.VBProject.VBComponents.import ThisWorkbook.Path & "\src\" & file
        file = Dir
  Wend
End Sub

