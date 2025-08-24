VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupingExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Expr
public expression as Expr
public function init(byval a as Expr) 
    nt st (expression, a)
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitGroupingExpr(Me)
    exit function
end function


public function Expr_copy() as variant
    dim  temp as GroupingExpr: st temp, new GroupingExpr
    call temp.init(expression.copy())
    st Expr_copy, temp
    exit function
end function



