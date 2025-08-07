Attribute VB_Name = "CallExpr"
implements Expr
public callee as Expr
public paren as Token
public arguments as variant
public function init(byval a as Expr, byval b as Token, byval c as variant) 
    nt st (callee, a)
    nt st (paren, b)
    nt st (arguments, c)
end function


private function Expr_copy() as variant
    dim  temp as CallExpr: st temp, new CallExpr
    dim  newargs() as variant
    dim  arg as Expr
    dim  i as long
    for i = 0 to al(arguments) - 1
        st arg, arguments(i)
        call arad(newargs, arg.copy())
        looplabel1bef:
    next i
    looplabel1aft:
    call temp.init(callee.copy(), paren, newargs)
    st Expr_copy, temp
    exit function
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitCallExpr(Me)
    exit function
end function



