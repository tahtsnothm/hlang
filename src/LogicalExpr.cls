VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogicalExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Expr
Public left As Expr
Public operator As Token
Public right As Expr

Public Function Expr_init(ParamArray arr() As Variant)
    st left, arr(0)
    st operator, arr(1)
    st right, arr(2)
End Function

Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitLogicalExpr(Me)
End Function

Private Function Expr_copy() As Variant
Dim temp As New LogicalExpr
temp.Expr_init left.copy(), operator, right.copy()
st Expr_copy, temp
End Function

