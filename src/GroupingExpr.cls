VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupingExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Expr
Public expression As Expr


Public Function Expr_init(ParamArray arr() As Variant)
    st expression, arr(0)

End Function

Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitGroupingExpr(Me)
End Function
Private Function Expr_copy() As Variant
Dim temp As New GroupingExpr
temp.Expr_init expression.copy()
st Expr_copy, temp
End Function
