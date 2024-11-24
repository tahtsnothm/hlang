VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Expr
Public callee As Expr
Public paren As Token
Public arguments As Variant ' list of expr

Public Function Expr_init(ParamArray arr() As Variant)
    st callee, arr(0)
    st paren, arr(1)
    st arguments, arr(2)
End Function

Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitCallExpr(Me)
End Function


