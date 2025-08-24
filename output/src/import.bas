Public Sub import()
   'use only on a blank workbook
   Dim MyObj As Object, MySource As Object, file As Variant
   file = dir(ThisWorkbook.Path & "\src\")
   Dim vbComp As Object
    Dim compsToRemove As Collection
    Set compsToRemove = New Collection
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
    If (vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule) And s <> "import" Then
            compsToRemove.Add vbComp
    End If
    Next vbComp
    Dim comp As Object
    For Each comp In compsToRemove
        ThisWorkbook.VBProject.VBComponents.Remove comp
    Next comp

   While (file <> "")
        Debug.Print "Import:", file
        
        ThisWorkbook.VBProject.VBComponents.import ThisWorkbook.Path & "\src\" & file
        file = dir
  Wend
End Sub

