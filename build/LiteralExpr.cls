VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteralExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Expr
Public value As Variant


Public Function Expr_init(ParamArray arr() As Variant)
    value = arr(0)

End Function

Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitLiteralExpr(Me)
End Function

Private Function Expr_copy() As Variant

st Expr_copy, Me
End Function
