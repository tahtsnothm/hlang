Attribute VB_Name = "LogicalExpr"
implements Expr
public left as Expr
public operator as Token
public right as Expr
public function init(byval a as Expr, byval b as Token, byval c as Expr) 
    nt st (left, a)
    dim  b as operator
    nt st (right, c)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitLogicalExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    dim  temp as LogicalExpr: st temp, new LogicalExpr
    call temp.init(left, operator, right)
    st Expr_copy, temp
    exit function
end function



