VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public name As Token
Public types_ As Variant
Public Function Stmt_init(ParamArray arr() As Variant)
    nt st(name, arr(0))
    nt st(types_, arr(1))
End Function


Public Function Stmt_accept(visitor As StmtVisitor)
    st Stmt_accept, visitor.visitEnumStmt(Me)
    Exit Function
End Function


