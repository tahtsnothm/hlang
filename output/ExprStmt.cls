Attribute VB_Name = "ExprStmt"
implements Stmt
public expression as Expr
public function init(byval a as Expr) 
    nt st (expression, a)
end function


public function Stmt_accept(byval visitor as StmtVisitor) as variant
    st Stmt_accept, visitor.visitExprStmt(Me)
    exit function
end function



