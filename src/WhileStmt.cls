VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhileStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public condition As Expr
Public body As Stmt
Public itx As Stmt


Public Function Stmt_init(ParamArray arr() As Variant)
    st condition, arr(0)
    st body, arr(1)
    st itx, arr(2)
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitWhileStmt(Me)
End Function



