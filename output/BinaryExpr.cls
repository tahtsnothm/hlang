Attribute VB_Name = "BinaryExpr"
implements Expr
dim  left as Expr
dim  operator as Token
dim  right as Expr
public function init(byval a as Expr, byval b as Token, byval c as Expr) 
    nt st (left, a)
    nt st (operator, b)
    nt st (right, c)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitBinaryExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    dim  temp as BinaryExpr: st temp, new BinaryExpr
    call temp.init(left.copy(), operator, right.copy())
    st Expr_copy, temp
    exit function
end function



