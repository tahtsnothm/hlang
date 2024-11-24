VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssignExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Expr
Public assignee As Expr
Public value As Expr

Public Function Expr_init(ParamArray arr() As Variant)
    st assignee, arr(0)
    st value, arr(1)
End Function

Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitAssignExpr(Me)
End Function


