VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AssignExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
implements Expr
public assignee as Expr
public value as Expr
public addnew as boolean
public function init(byval a as Expr, byval b as Expr, byval c as boolean) 
    nt st (assignee, a)
    nt st (value, b)
    nt st (addnew, c)
end function


private function Expr_copy() as variant
    dim  temp as AssignExpr: st temp, new AssignExpr
    call temp.init(assignee.copy(), value.copy(), addnew)
    st Expr_copy, temp
    exit function
end function


public function Expr_accept(byval visitor as ExprVisitor) as variant
    st Expr_accept, visitor.visitAssignExpr(Me)
    exit function
end function



