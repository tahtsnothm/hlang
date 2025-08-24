VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BinaryExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



