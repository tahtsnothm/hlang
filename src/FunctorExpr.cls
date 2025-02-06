VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctorExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements Expr
Public name As String
Public Function Expr_init(ParamArray arr() As Variant)
    nt st(name, arr(0))
End Function


Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitFunctorExpr(Me)
    Exit Function
End Function


Public Function Expr_copy() As Variant
    Dim temp As FunctorExpr: st temp, New FunctorExpr
    Call temp.Expr_init(name)
    st Expr_copy, temp
    Exit Function
End Function




