Attribute VB_Name = "import"
Public Sub import()
   'use only on a blank workbook
   Dim MyObj As Object, MySource As Object, file As Variant
   file = Dir(ThisWorkbook.Path & "\src\")
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
    Next vbComp
   While (file <> "")
        pt "Import:", file
        ThisWorkbook.VBProject.VBComponents.import ThisWorkbook.Path & "\src\" & file
        file = Dir
  Wend
End Sub
