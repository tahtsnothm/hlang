Attribute VB_Name = "FunctExpr"
implements Expr
public name as string
public function init(byval a as string) 
    nt st (name, a)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitFunctExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    dim  temp as FunctExpr: st temp, new FunctExpr
    call temp.init(name)
    st Expr_copy, temp
    exit function
end function



