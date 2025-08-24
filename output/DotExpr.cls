VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



