Attribute VB_Name = "UnaryExpr"
implements Expr
public operator as Token
public right as Expr
public function init(byval a as Token, byval b as Expr) 
    nt st (operator, a)
    nt st (right, b)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitUnaryExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    dim  temp as UnaryExpr: st temp, new UnaryExpr
    call temp.init(operator, right)
    st Expr_copy, temp
    exit function
end function



