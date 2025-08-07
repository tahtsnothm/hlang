Attribute VB_Name = "LiteralExpr"
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



