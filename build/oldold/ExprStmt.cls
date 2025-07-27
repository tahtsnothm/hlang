VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExprStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public expression As Expr

Public Function Stmt_init(ParamArray arr() As Variant)
    st expression, arr(0)

End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitExprStmt(Me)
End Function

