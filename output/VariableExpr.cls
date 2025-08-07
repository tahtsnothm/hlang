Attribute VB_Name = "VariableExpr"
implements Expr
public name as Token
public function init(byval a as Token) 
    nt st (name, a)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitVariableExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    st Expr_copy, me
    exit function
end function



