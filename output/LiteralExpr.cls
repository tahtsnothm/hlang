VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LiteralExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Expr
public value as variant
public function init(byval a as variant) 
    nt st (value, a)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitLiteralExpr(Me)
    exit function
end function


public function Expr_copy() as variant
    st Expr_copy, me
    exit function
end function



