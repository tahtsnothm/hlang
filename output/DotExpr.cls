Attribute VB_Name = "DotExpr"
implements Expr
public left as Expr
public right as Expr
public function init(byval a as Expr, byval b as Expr) 
    nt st (left, a)
    nt st (right, b)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitDotExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    dim  temp as DotExpr: st temp, new DotExpr
    call temp.init(left.copy(), right.copy())
    st Expr_copy, temp
    exit function
end function



