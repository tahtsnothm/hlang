VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReturnStmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Stmt
Public keyword As Token
Public hasValue As Boolean

Public value As Expr


Public Function Stmt_init(ParamArray arr() As Variant)
    st keyword, arr(0)
    st value, arr(1)
    st hasValue, arr(2)
End Function

Public Function Stmt_accept(visitor As StmtVisitor) As Variant
    st Stmt_accept, visitor.visitReturnStmt(Me)
End Function




