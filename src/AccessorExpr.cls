VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessorExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Expr
Public accessee As Expr
Public paren As Token
Public arguments As Variant ' list of expr
Public valMode As String ' access, slice, val, access_set, val_set ( lvalue
Public isLvalue As Boolean

Public Function Expr_init(ParamArray arr() As Variant)
    st accessee, arr(0)
    st paren, arr(1)
    st arguments, arr(2)
    st valMode, arr(3)
    st isLvalue, False
End Function

Public Function Expr_accept(visitor As ExprVisitor) As Variant
    st Expr_accept, visitor.visitAccessorExpr(Me)
End Function

Private Function Expr_copy() As Variant
Dim temp As New AccessorExpr
''' replace
Dim newargs() As Variant
Dim arg As Expr
Dim i As Long
For i = 0 To al(arguments) - 1
    st arg, arguments(i)
    arad newargs, arg.copy()
Next i
temp.Expr_init accessee.copy(), paren, newargs, valMode
st Expr_copy, temp
End Function

