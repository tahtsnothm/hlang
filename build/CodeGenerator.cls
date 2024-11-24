VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements ExprVisitor
Implements StmtVisitor
Public output_file As String
Public functions As String

Private Sub Class_Initialize()
    ' default function will be in main
    ' theres a main thread for stuff
    ' and another one for functions
    output_file = ThisWorkbook.path & "\out.bas"
    ' to change when classes are supported
End Sub
Public Function generate(statements As Variant)
    Dim statement As Stmt
    Dim i As Long
    For i = 0 To al(statements) - 1
        st statement, statements(i)
        statement.accept (Me)
    Next i
End Function


