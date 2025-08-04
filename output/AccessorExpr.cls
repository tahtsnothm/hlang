Attribute VB_Name = "AccessorExpr"
implements Expr
public accessee as Expr
public paren as Token
public arguments as variant
public valMode as string
public isLvalue as boolean
public function init(byval a as Expr, byval b as Token, byval c as variant, byval d as string) 
    nt st (accessee, a)
    nt st (paren, b)
    nt st (arguments, c)
    nt st (valMode, d)
    nt st (isLvalue, False)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitAccessorExpr(Me)
    exit function
end function


private function Expr_copy() as variant
    dim  temp as AccessorExpr: st temp, new AccessorExpr
    dim  newargs() as variant
    dim  arg as Expr
    dim  i as long
    for i = 0 to al(arguments) - 1
        st arg, arguments(i)
        call arad(newargs, arg.copy())
        looplabel0bef:
    next i
    looplabel0aft:
    call temp.init(accessee.copy(), paren, newargs, valMode)
    st Expr_copy, temp
    exit function
end function



